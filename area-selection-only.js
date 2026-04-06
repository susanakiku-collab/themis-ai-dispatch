// THEMIS AI Dispatch
// 方面選定 + 残りキャストの入れ込みだけを独立させた関数
// 元になっている考え方:
// - dispatch-engine.js の selectPrimaryAxes
// - dispatch-engine.js の runInitialDirectionalDispatch
// - ただし再分散系ロジックは入れず、最初のグループ作成だけに限定

const AREA_SELECTION_SEPARATION_THRESHOLD = 70;
const SAME_DIRECTION_THRESHOLD = 45;
const NATURAL_TURN_THRESHOLD = 30;
const PRIMARY_ANGLE_WEIGHT = 0.30;

function normalizeAngle(angle) {
  const normalized = Number(angle || 0) % 360;
  return normalized < 0 ? normalized + 360 : normalized;
}

function angleDiff(a, b) {
  const diff = Math.abs(normalizeAngle(a) - normalizeAngle(b));
  return diff > 180 ? 360 - diff : diff;
}

function signedAngleDiff(fromAngle, toAngle) {
  let diff = normalizeAngle(toAngle) - normalizeAngle(fromAngle);
  if (diff > 180) diff -= 360;
  if (diff < -180) diff += 360;
  return diff;
}

function calcAngleFromOrigin(origin, point) {
  const dx = Number(point.lng) - Number(origin.lng);
  const dy = Number(point.lat) - Number(origin.lat);
  return normalizeAngle((Math.atan2(dy, dx) * 180) / Math.PI);
}

function distanceKm(a, b) {
  const lat1 = Number(a.lat);
  const lng1 = Number(a.lng);
  const lat2 = Number(b.lat);
  const lng2 = Number(b.lng);

  if (![lat1, lng1, lat2, lng2].every(Number.isFinite)) return 0;

  const toRad = deg => (deg * Math.PI) / 180;
  const R = 6371;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const p1 = toRad(lat1);
  const p2 = toRad(lat2);

  const x = Math.sin(dLat / 2) ** 2 + Math.cos(p1) * Math.cos(p2) * Math.sin(dLng / 2) ** 2;
  const y = 2 * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
  return R * y;
}

function enrichPersonForAreaSelection(origin, person) {
  return {
    ...person,
    area: String(person.area || '').trim(),
    angleFromOrigin: calcAngleFromOrigin(origin, person),
    distanceFromOrigin: distanceKm(origin, person)
  };
}

function isFarEnoughFromSeeds(candidateAngle, seeds, threshold = AREA_SELECTION_SEPARATION_THRESHOLD) {
  if (!Array.isArray(seeds) || !seeds.length) return true;
  return seeds.every(seed => angleDiff(candidateAngle, seed.angle) >= threshold);
}

function assignPeopleToClosestSeed(people, seeds) {
  return seeds.map(seed => ({
    seedId: seed.seedId,
    angle: seed.angle,
    members: people.filter(person => {
      let bestSeed = null;
      let bestDiff = Infinity;

      seeds.forEach(candidate => {
        const diff = angleDiff(person.angleFromOrigin, candidate.angle);
        if (diff < bestDiff) {
          bestDiff = diff;
          bestSeed = candidate;
        }
      });

      return bestSeed && bestSeed.seedId === seed.seedId;
    })
  }));
}

function chooseAdditionalSeedFromExistingGroups(groups, selectedSeeds) {
  const usedIds = new Set((selectedSeeds || []).map(seed => String(seed.sourcePersonId || '')));

  const sortedGroups = [...(groups || [])].sort((a, b) => {
    if ((b.members?.length || 0) !== (a.members?.length || 0)) {
      return (b.members?.length || 0) - (a.members?.length || 0);
    }
    const aMax = (a.members || []).reduce((max, p) => Math.max(max, Number(p.distanceFromOrigin || 0)), 0);
    const bMax = (b.members || []).reduce((max, p) => Math.max(max, Number(p.distanceFromOrigin || 0)), 0);
    return bMax - aMax;
  });

  for (const group of sortedGroups) {
    const candidate = [...(group.members || [])]
      .filter(person => !usedIds.has(String(person.id || '')))
      .sort((a, b) => Number(b.distanceFromOrigin || 0) - Number(a.distanceFromOrigin || 0))[0];

    if (candidate) return candidate;
  }

  return null;
}

function selectPrimaryAreasByFarthest(origin, people, vehicleCount, options = {}) {
  const separationThreshold = Number(options.separationThreshold || AREA_SELECTION_SEPARATION_THRESHOLD);
  const maxCount = Math.max(0, Number(vehicleCount || 0));

  const enrichedPeople = (Array.isArray(people) ? people : [])
    .filter(person => person && Number.isFinite(Number(person.lat)) && Number.isFinite(Number(person.lng)))
    .map(person => enrichPersonForAreaSelection(origin, person));

  const seeds = [];
  if (!enrichedPeople.length || maxCount <= 0) {
    return { selectedAreas: [], selectedSeeds: [], people: enrichedPeople };
  }

  const remaining = [...enrichedPeople].sort((a, b) => {
    if (Number(b.distanceFromOrigin || 0) !== Number(a.distanceFromOrigin || 0)) {
      return Number(b.distanceFromOrigin || 0) - Number(a.distanceFromOrigin || 0);
    }
    return String(a.id || '').localeCompare(String(b.id || ''), 'ja');
  });

  const first = remaining.shift();
  seeds.push({
    seedId: 'seed_1',
    area: first.area,
    angle: first.angleFromOrigin,
    distanceFromOrigin: first.distanceFromOrigin,
    sourcePersonId: first.id,
    sourcePersonName: first.name || ''
  });

  while (seeds.length < maxCount) {
    const index = remaining.findIndex(person => isFarEnoughFromSeeds(person.angleFromOrigin, seeds, separationThreshold));
    if (index === -1) break;

    const person = remaining.splice(index, 1)[0];
    seeds.push({
      seedId: `seed_${seeds.length + 1}`,
      area: person.area,
      angle: person.angleFromOrigin,
      distanceFromOrigin: person.distanceFromOrigin,
      sourcePersonId: person.id,
      sourcePersonName: person.name || ''
    });
  }

  while (seeds.length < maxCount) {
    const groups = assignPeopleToClosestSeed(enrichedPeople, seeds);
    const extra = chooseAdditionalSeedFromExistingGroups(groups, seeds);
    if (!extra) break;

    seeds.push({
      seedId: `seed_${seeds.length + 1}`,
      area: extra.area,
      angle: extra.angleFromOrigin,
      distanceFromOrigin: extra.distanceFromOrigin,
      sourcePersonId: extra.id,
      sourcePersonName: extra.name || ''
    });
  }

  const seenAreas = new Set();
  const selectedAreas = [];
  seeds.forEach(seed => {
    const area = String(seed.area || '').trim();
    if (!area) return;
    if (seenAreas.has(area)) return;
    seenAreas.add(area);
    selectedAreas.push(area);
  });

  return {
    selectedAreas,
    selectedSeeds: seeds,
    people: enrichedPeople
  };
}

function attachClosestSeedToPeople(people, seeds) {
  return (Array.isArray(people) ? people : []).map(person => {
    let bestSeed = null;
    let bestDiff = Infinity;

    (Array.isArray(seeds) ? seeds : []).forEach(seed => {
      const diff = angleDiff(person.angleFromOrigin, seed.angle);
      if (diff < bestDiff) {
        bestDiff = diff;
        bestSeed = seed;
      }
    });

    return {
      ...person,
      closestSeedId: bestSeed ? bestSeed.seedId : null,
      closestSeedAngle: bestSeed ? bestSeed.angle : null,
      closestSeedArea: bestSeed ? bestSeed.area : '',
      seedAngleDiff: bestDiff === Infinity ? null : bestDiff
    };
  });
}

function createInitialAreaGroups(vehicles = [], seeds = []) {
  const count = Math.max(
    Array.isArray(vehicles) ? vehicles.length : 0,
    Array.isArray(seeds) ? seeds.length : 0
  );

  return Array.from({ length: count }, (_, index) => ({
    groupId: `group_${index + 1}`,
    vehicleId: vehicles[index]?.id ?? null,
    vehicleName: vehicles[index]?.name || vehicles[index]?.vehicleName || '',
    capacity: Number(vehicles[index]?.capacity || vehicles[index]?.seat_capacity || Number.MAX_SAFE_INTEGER),
    assigned: [],
    primaryAngle: null,
    seedId: seeds[index]?.seedId || null,
    seedAngle: seeds[index]?.angle ?? null,
    seedArea: seeds[index]?.area || '',
    seedSourcePersonId: seeds[index]?.sourcePersonId || null,
    seedSourcePersonName: seeds[index]?.sourcePersonName || ''
  }));
}

function updatePrimaryAngleSoft(oldPrimaryAngle, personAngle) {
  if (oldPrimaryAngle == null) return normalizeAngle(personAngle);
  const diff = signedAngleDiff(oldPrimaryAngle, personAngle);
  return normalizeAngle(oldPrimaryAngle + diff * PRIMARY_ANGLE_WEIGHT);
}

function isSameDirection(primaryAngle, personAngle) {
  if (primaryAngle == null || personAngle == null) return false;
  return angleDiff(primaryAngle, personAngle) <= SAME_DIRECTION_THRESHOLD;
}

function isReverseDirection(origin, lastPerson, person) {
  if (!origin || !lastPerson || !person) return false;

  const ax = Number(lastPerson.lng) - Number(origin.lng);
  const ay = Number(lastPerson.lat) - Number(origin.lat);
  const bx = Number(person.lng) - Number(origin.lng);
  const by = Number(person.lat) - Number(origin.lat);

  const abx = bx - ax;
  const aby = by - ay;
  const abLen2 = abx * abx + aby * aby;
  if (abLen2 === 0) return false;

  const t = -((ax * abx) + (ay * aby)) / abLen2;
  if (t <= 0 || t >= 1) return false;

  const px = ax + abx * t;
  const py = ay + aby * t;

  const pDist2 = px * px + py * py;
  const aDist2 = ax * ax + ay * ay;
  const bDist2 = bx * bx + by * by;

  return pDist2 < Math.min(aDist2, bDist2);
}

function seedPriorityForGroup(group, person) {
  if (group.seedId && person.closestSeedId && group.seedId === person.closestSeedId) return 0;
  if ((group.assigned || []).length === 0) return 1;
  return 2;
}

function calcRouteDistance(origin, assigned) {
  const rows = Array.isArray(assigned) ? assigned : [];
  if (!rows.length) return 0;

  let total = 0;
  let current = origin;
  rows.forEach(person => {
    total += distanceKm(current, person);
    current = person;
  });
  return Number(total.toFixed(3));
}

function calcReturnDistance(origin, assigned) {
  const rows = Array.isArray(assigned) ? assigned : [];
  if (!rows.length) return 0;
  return Number(distanceKm(rows[rows.length - 1], origin).toFixed(3));
}

function sortAssignedByOriginDistance(assigned) {
  return [...(Array.isArray(assigned) ? assigned : [])].sort((a, b) => {
    if (Number(a.distanceFromOrigin || 0) !== Number(b.distanceFromOrigin || 0)) {
      return Number(a.distanceFromOrigin || 0) - Number(b.distanceFromOrigin || 0);
    }
    return String(a.id || '').localeCompare(String(b.id || ''), 'ja');
  });
}

function buildTrialCandidate(origin, group, person) {
  const isEmptyGroup = (group.assigned || []).length === 0;
  const seedPriority = seedPriorityForGroup(group, person);

  if (isEmptyGroup) {
    const trialAssigned = [person];
    return {
      groupId: group.groupId,
      isEmptyGroup: true,
      seedPriority,
      reverseFlag: false,
      sameDirection: null,
      turnAmount: null,
      routeDistance: calcRouteDistance(origin, trialAssigned),
      returnDistance: calcReturnDistance(origin, trialAssigned),
      totalDistance: calcRouteDistance(origin, trialAssigned) + calcReturnDistance(origin, trialAssigned),
      assignedCount: 1,
      newPrimaryAngle: person.angleFromOrigin
    };
  }

  const oldPrimaryAngle = group.primaryAngle != null ? group.primaryAngle : group.seedAngle;
  const newPrimaryAngle = updatePrimaryAngleSoft(oldPrimaryAngle, person.angleFromOrigin);
  const lastPerson = group.assigned[group.assigned.length - 1];
  const trialAssigned = [...group.assigned, person];

  return {
    groupId: group.groupId,
    isEmptyGroup: false,
    seedPriority,
    reverseFlag: isReverseDirection(origin, lastPerson, person),
    sameDirection: isSameDirection(oldPrimaryAngle, person.angleFromOrigin),
    turnAmount: angleDiff(oldPrimaryAngle, newPrimaryAngle),
    routeDistance: calcRouteDistance(origin, trialAssigned),
    returnDistance: calcReturnDistance(origin, trialAssigned),
    totalDistance: calcRouteDistance(origin, trialAssigned) + calcReturnDistance(origin, trialAssigned),
    assignedCount: trialAssigned.length,
    newPrimaryAngle
  };
}

function turnBand(turnAmount) {
  if (turnAmount == null) return 99;
  if (turnAmount <= 15) return 0;
  if (turnAmount <= 30) return 1;
  if (turnAmount <= 45) return 2;
  return 3;
}

function candidateGroup(candidate) {
  if (
    candidate.isEmptyGroup === false &&
    candidate.reverseFlag === false &&
    candidate.sameDirection === true &&
    candidate.turnAmount != null &&
    candidate.turnAmount <= NATURAL_TURN_THRESHOLD
  ) {
    return 0;
  }
  if (candidate.isEmptyGroup === true) return 1;
  return 2;
}

function compareTrialCandidate(a, b) {
  if (a.seedPriority !== b.seedPriority) return a.seedPriority - b.seedPriority;

  const groupA = candidateGroup(a);
  const groupB = candidateGroup(b);
  if (groupA !== groupB) return groupA - groupB;

  if (a.reverseFlag !== b.reverseFlag) return a.reverseFlag ? 1 : -1;

  if (a.sameDirection !== b.sameDirection) {
    if (a.sameDirection === true) return -1;
    if (b.sameDirection === true) return 1;
  }

  const bandA = turnBand(a.turnAmount);
  const bandB = turnBand(b.turnAmount);
  if (bandA !== bandB) return bandA - bandB;

  if (a.turnAmount != null && b.turnAmount != null && a.turnAmount !== b.turnAmount) {
    return a.turnAmount - b.turnAmount;
  }

  if (a.returnDistance !== b.returnDistance) return a.returnDistance - b.returnDistance;
  if (a.totalDistance !== b.totalDistance) return a.totalDistance - b.totalDistance;
  if (a.assignedCount !== b.assignedCount) return a.assignedCount - b.assignedCount;

  return String(a.groupId || '').localeCompare(String(b.groupId || ''), 'ja');
}

function findBestGroupAssignment(origin, groups, person) {
  const candidates = [];

  (Array.isArray(groups) ? groups : []).forEach((group, groupIndex) => {
    const capacity = Number(group.capacity || Number.MAX_SAFE_INTEGER);
    if ((group.assigned || []).length >= capacity) return;

    candidates.push({
      groupIndex,
      candidate: buildTrialCandidate(origin, group, person)
    });
  });

  if (!candidates.length) return null;
  candidates.sort((a, b) => compareTrialCandidate(a.candidate, b.candidate));
  return candidates[0];
}

function assignRemainingPeopleIntoSelectedAreas(origin, groups, people) {
  const workingGroups = Array.isArray(groups)
    ? groups.map(group => ({ ...group, assigned: [...(group.assigned || [])] }))
    : [];
  const overflow = [];

  const sortedPeople = sortAssignedByOriginDistance(people);

  sortedPeople.forEach(person => {
    const best = findBestGroupAssignment(origin, workingGroups, person);
    if (!best) {
      overflow.push(person);
      return;
    }

    const group = workingGroups[best.groupIndex];
    group.assigned.push(person);
    group.primaryAngle = best.candidate.newPrimaryAngle;
    if (!group.seedId && person.closestSeedId) {
      group.seedId = person.closestSeedId;
      group.seedAngle = person.closestSeedAngle;
      group.seedArea = person.closestSeedArea || group.seedArea;
    }
  });

  workingGroups.forEach(group => {
    group.assigned = sortAssignedByOriginDistance(group.assigned);
    if ((group.assigned || []).length) {
      let angle = group.seedAngle != null ? group.seedAngle : group.assigned[0].angleFromOrigin;
      group.assigned.forEach(person => {
        angle = updatePrimaryAngleSoft(angle, person.angleFromOrigin);
      });
      group.primaryAngle = normalizeAngle(angle);
    } else {
      group.primaryAngle = group.seedAngle != null ? normalizeAngle(group.seedAngle) : null;
    }

    group.routeDistance = calcRouteDistance(origin, group.assigned);
    group.returnDistance = calcReturnDistance(origin, group.assigned);
    group.totalDistance = Number((group.routeDistance + group.returnDistance).toFixed(3));
  });

  return {
    groups: workingGroups,
    overflow
  };
}

function buildPrimaryAreaGroups(origin, vehicles, people, options = {}) {
  const vehicleCount = Math.max(0, Number(options.vehicleCount || vehicles?.length || 0));
  const selected = selectPrimaryAreasByFarthest(origin, people, vehicleCount, options);
  const peopleWithSeed = attachClosestSeedToPeople(selected.people, selected.selectedSeeds);
  const groups = createInitialAreaGroups(Array.isArray(vehicles) ? vehicles : [], selected.selectedSeeds);
  const assigned = assignRemainingPeopleIntoSelectedAreas(origin, groups, peopleWithSeed);

  return {
    selectedAreas: selected.selectedAreas,
    selectedSeeds: selected.selectedSeeds,
    people: peopleWithSeed,
    groups: assigned.groups,
    overflow: assigned.overflow
  };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    selectPrimaryAreasByFarthest,
    buildPrimaryAreaGroups,
    assignRemainingPeopleIntoSelectedAreas,
    attachClosestSeedToPeople,
    createInitialAreaGroups,
    normalizeAngle,
    angleDiff,
    signedAngleDiff,
    calcAngleFromOrigin,
    distanceKm,
    enrichPersonForAreaSelection,
    isFarEnoughFromSeeds,
    assignPeopleToClosestSeed,
    chooseAdditionalSeedFromExistingGroups,
    updatePrimaryAngleSoft,
    isSameDirection,
    isReverseDirection,
    findBestGroupAssignment,
    buildTrialCandidate
  };
}

if (typeof window !== 'undefined') {
  window.selectPrimaryAreasByFarthest = selectPrimaryAreasByFarthest;
  window.buildPrimaryAreaGroups = buildPrimaryAreaGroups;
  window.assignRemainingPeopleIntoSelectedAreas = assignRemainingPeopleIntoSelectedAreas;
}
