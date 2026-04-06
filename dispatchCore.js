window.DispatchCore = (function () {
  const ANGLE_THRESHOLD = 70;
  const EPS = 1e-9;

  function toNumber(value, fallback = 0) {
    const n = Number(value);
    return Number.isFinite(n) ? n : fallback;
  }

  function normalizeAngle(angle) {
    const n = toNumber(angle, 0) % 360;
    return n < 0 ? n + 360 : n;
  }

  function angleDiff(a, b) {
    const diff = Math.abs(normalizeAngle(a) - normalizeAngle(b));
    return diff > 180 ? 360 - diff : diff;
  }

  function getOrigin() {
    const lat = toNumber(globalThis.ORIGIN_LAT ?? globalThis.window?.APP_CONFIG?.ORIGIN_LAT, NaN);
    const lng = toNumber(globalThis.ORIGIN_LNG ?? globalThis.window?.APP_CONFIG?.ORIGIN_LNG, NaN);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
    return { lat, lng };
  }

  function distanceKm(a, b) {
    const lat1 = toNumber(a?.lat, NaN);
    const lng1 = toNumber(a?.lng, NaN);
    const lat2 = toNumber(b?.lat, NaN);
    const lng2 = toNumber(b?.lng, NaN);
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

  function calcAngleFromOrigin(origin, point) {
    const dx = toNumber(point?.lng, 0) - toNumber(origin?.lng, 0);
    const dy = toNumber(point?.lat, 0) - toNumber(origin?.lat, 0);
    return normalizeAngle((Math.atan2(dy, dx) * 180) / Math.PI);
  }

  function normalizeAreaLabel(value) {
    if (typeof globalThis.normalizeAreaLabel === 'function') {
      return globalThis.normalizeAreaLabel(value || '');
    }
    return String(value || '').trim();
  }

  function normalizeVehicle(vehicle) {
    return {
      id: toNumber(vehicle?.id, 0),
      driver_name: String(vehicle?.driver_name || vehicle?.name || vehicle?.vehicleName || vehicle?.plate_number || ''),
      seat_capacity: Math.max(1, toNumber(vehicle?.seat_capacity ?? vehicle?.capacity, 4)),
      raw: vehicle
    };
  }

  function normalizePerson(item, origin) {
    const lat = toNumber(item?.casts?.latitude ?? item?.latitude ?? item?.lat, NaN);
    const lng = toNumber(item?.casts?.longitude ?? item?.longitude ?? item?.lng, NaN);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;

    const point = { lat, lng };
    const angleFromOrigin = calcAngleFromOrigin(origin, point);
    const distanceFromOrigin = distanceKm(origin, point);
    const dispatchDistance = toNumber(item?.distance_km ?? item?.casts?.distance_km, 0);

    return {
      id: toNumber(item?.id, 0),
      cast_id: toNumber(item?.cast_id, 0),
      name: String(item?.casts?.name || item?.name || ''),
      actual_hour: toNumber(item?.actual_hour ?? item?.plan_hour, 0),
      lat,
      lng,
      angleFromOrigin,
      distanceFromOrigin,
      distance_km: dispatchDistance,
      priorityDistance: dispatchDistance > 0 ? dispatchDistance : distanceFromOrigin,
      destination_area: normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || ''),
      raw: item
    };
  }

  function sortByAnchorPriorityDesc(people) {
    return [...people].sort((a, b) => {
      if (Math.abs((b.priorityDistance || 0) - (a.priorityDistance || 0)) > EPS) {
        return (b.priorityDistance || 0) - (a.priorityDistance || 0);
      }
      if (Math.abs((b.distanceFromOrigin || 0) - (a.distanceFromOrigin || 0)) > EPS) {
        return (b.distanceFromOrigin || 0) - (a.distanceFromOrigin || 0);
      }
      return String(a.id || '').localeCompare(String(b.id || ''), 'ja');
    });
  }

  function sortByOriginDistanceAsc(people) {
    return [...people].sort((a, b) => {
      if (Math.abs((a.distanceFromOrigin || 0) - (b.distanceFromOrigin || 0)) > EPS) {
        return (a.distanceFromOrigin || 0) - (b.distanceFromOrigin || 0);
      }
      if (Math.abs((a.priorityDistance || 0) - (b.priorityDistance || 0)) > EPS) {
        return (a.priorityDistance || 0) - (b.priorityDistance || 0);
      }
      return String(a.id || '').localeCompare(String(b.id || ''), 'ja');
    });
  }

  function minAngleDiffToAnchors(person, anchors) {
    if (!Array.isArray(anchors) || !anchors.length) return 999;
    return anchors.reduce((min, anchor) => {
      return Math.min(min, angleDiff(person.angleFromOrigin, anchor.angleFromOrigin));
    }, 999);
  }

  function selectAnchors(people, vehicleCount) {
    const maxCount = Math.max(0, toNumber(vehicleCount, 0));
    if (!Array.isArray(people) || !people.length || maxCount <= 0) return [];

    const sorted = sortByAnchorPriorityDesc(people);
    const anchors = [];
    const usedIds = new Set();

    const first = sorted[0];
    if (first) {
      anchors.push(first);
      usedIds.add(String(first.id));
    }

    while (anchors.length < maxCount) {
      const remaining = sorted.filter(person => !usedIds.has(String(person.id)));
      if (!remaining.length) break;

      const outside = remaining.filter(person => {
        return anchors.every(anchor => angleDiff(person.angleFromOrigin, anchor.angleFromOrigin) >= ANGLE_THRESHOLD);
      });

      const pool = outside.length ? outside : remaining;
      const picked = [...pool].sort((a, b) => {
        if (Math.abs((b.priorityDistance || 0) - (a.priorityDistance || 0)) > EPS) {
          return (b.priorityDistance || 0) - (a.priorityDistance || 0);
        }
        const aMin = minAngleDiffToAnchors(a, anchors);
        const bMin = minAngleDiffToAnchors(b, anchors);
        if (Math.abs(bMin - aMin) > EPS) return bMin - aMin;
        if (Math.abs((b.distanceFromOrigin || 0) - (a.distanceFromOrigin || 0)) > EPS) {
          return (b.distanceFromOrigin || 0) - (a.distanceFromOrigin || 0);
        }
        return String(a.id || '').localeCompare(String(b.id || ''), 'ja');
      })[0];

      if (!picked) break;
      anchors.push(picked);
      usedIds.add(String(picked.id));
    }

    return anchors.slice(0, maxCount);
  }

  function createGroups(hourVehicles, anchors) {
    return anchors.map((anchor, index) => {
      const vehicle = hourVehicles[index] || {};
      return {
        groupId: `group_${index + 1}`,
        vehicleId: toNumber(vehicle.id, 0),
        vehicleName: String(vehicle.driver_name || vehicle.name || ''),
        capacity: Math.max(1, toNumber(vehicle.seat_capacity ?? vehicle.capacity, 4)),
        anchor,
        anchorId: toNumber(anchor.id, 0),
        members: [],
        routeOrder: [anchor],
        routeDistance: 0,
        totalDistance: 0
      };
    });
  }

  function chooseTargetGroup(groups, person) {
    const candidates = (groups || []).map((group, index) => ({
      group,
      index,
      diff: angleDiff(person.angleFromOrigin, group.anchor?.angleFromOrigin),
      count: Array.isArray(group.members) ? group.members.length : 0
    })).filter(entry => {
      const memberCount = Array.isArray(entry.group.members) ? entry.group.members.length : 0;
      return memberCount < Math.max(0, toNumber(entry.group.capacity, 0) - 1);
    });

    if (!candidates.length) return null;

    const within = candidates.filter(entry => entry.diff <= ANGLE_THRESHOLD);
    const pool = within.length ? within : candidates;

    pool.sort((a, b) => {
      if (Math.abs(a.diff - b.diff) > EPS) return a.diff - b.diff;
      const aAnchorDistance = toNumber(a.group.anchor?.priorityDistance, 0);
      const bAnchorDistance = toNumber(b.group.anchor?.priorityDistance, 0);
      if (Math.abs(bAnchorDistance - aAnchorDistance) > EPS) return bAnchorDistance - aAnchorDistance;
      if (a.count !== b.count) return a.count - b.count;
      return toNumber(a.group.vehicleId, 0) - toNumber(b.group.vehicleId, 0);
    });

    return pool[0];
  }

  function buildRouteOrder(group) {
    const sortedMembers = sortByOriginDistanceAsc(group.members || []);
    return [...sortedMembers, group.anchor].filter(Boolean);
  }

  function calcRouteDistance(origin, orderedPeople) {
    const rows = Array.isArray(orderedPeople) ? orderedPeople.filter(Boolean) : [];
    if (!rows.length) return 0;
    let total = 0;
    let current = origin;
    rows.forEach(person => {
      total += distanceKm(current, person);
      current = person;
    });
    return Number(total.toFixed(6));
  }

  function refreshGroupMetrics(origin, groups) {
    (groups || []).forEach(group => {
      group.members = sortByOriginDistanceAsc(group.members || []);
      group.routeOrder = buildRouteOrder(group);
      group.routeDistance = calcRouteDistance(origin, group.routeOrder);
      group.totalDistance = group.routeDistance;
    });
  }

  function buildAssignmentsFromGroups(groups, hour) {
    const rows = [];
    (groups || []).forEach(group => {
      const ordered = Array.isArray(group.routeOrder) ? group.routeOrder : [];
      ordered.forEach((person, index) => {
        rows.push({
          item_id: toNumber(person.id, 0),
          vehicle_id: toNumber(group.vehicleId, 0),
          actual_hour: toNumber(hour, 0),
          driver_name: String(group.vehicleName || ''),
          distance_km: toNumber(person.raw?.distance_km ?? person.distance_km, 0),
          stop_order: index + 1
        });
      });
    });
    return rows;
  }

  function buildHourAssignments(origin, hourVehicles, hourPeople) {
    const anchors = selectAnchors(hourPeople, hourVehicles.length);
    const groups = createGroups(hourVehicles, anchors);
    const anchorIds = new Set(anchors.map(anchor => String(anchor.id)));
    const remaining = sortByOriginDistanceAsc((hourPeople || []).filter(person => !anchorIds.has(String(person.id))));
    const overflow = [];
    const evaluations = [];

    remaining.forEach(person => {
      const target = chooseTargetGroup(groups, person);
      if (!target) {
        overflow.push(person);
        evaluations.push({
          itemId: toNumber(person.id, 0),
          castName: String(person.name || ''),
          selectedVehicleId: 0,
          selectedGlobalDistance: 0,
          hour: toNumber(person.actual_hour, 0),
          reason: 'no_available_capacity'
        });
        return;
      }

      target.group.members.push(person);
      evaluations.push({
        itemId: toNumber(person.id, 0),
        castName: String(person.name || ''),
        selectedVehicleId: toNumber(target.group.vehicleId, 0),
        selectedGlobalDistance: 0,
        hour: toNumber(person.actual_hour, 0),
        anchorName: String(target.group.anchor?.name || ''),
        anchorId: toNumber(target.group.anchor?.id, 0),
        angleDiff: Number(target.diff || 0)
      });
    });

    refreshGroupMetrics(origin, groups);
    return { groups, anchors, overflow, evaluations };
  }

  function optimizeAssignments(items, vehicles, monthlyMap, options = {}) {
    const origin = getOrigin();
    const safeItems = (Array.isArray(items) ? items : []).filter(Boolean);
    const safeVehicles = (Array.isArray(vehicles) ? vehicles : []).filter(Boolean).map(normalizeVehicle).filter(v => v.id);

    if (!origin || !safeItems.length || !safeVehicles.length) {
      globalThis.__THEMIS_LAST_OVERFLOW__ = {
        mode: 'anchor_slot_fixed',
        vehicleCount: safeVehicles.length,
        actualGroupCount: 0,
        keptGroups: [],
        overflowGroups: [],
        evaluations: [],
        totalSeatCapacity: safeVehicles.reduce((sum, v) => sum + Math.max(1, toNumber(v.seat_capacity, 4)), 0),
        totalCastCount: safeItems.length,
        capacityOverflowCount: 0,
        capacityOverflowItems: [],
        selectedAreas: []
      };
      return [];
    }

    const validPeople = [];
    const invalidPeople = [];
    safeItems.forEach(item => {
      const person = normalizePerson(item, origin);
      if (person) validPeople.push(person);
      else invalidPeople.push(item);
    });

    if (!validPeople.length) {
      globalThis.__THEMIS_LAST_OVERFLOW__ = {
        mode: 'anchor_slot_fixed',
        vehicleCount: safeVehicles.length,
        actualGroupCount: 0,
        keptGroups: [],
        overflowGroups: [],
        evaluations: [],
        totalSeatCapacity: safeVehicles.reduce((sum, v) => sum + Math.max(1, toNumber(v.seat_capacity, 4)), 0),
        totalCastCount: safeItems.length,
        capacityOverflowCount: invalidPeople.length,
        capacityOverflowItems: invalidPeople.map(item => ({
          itemId: toNumber(item?.id, 0),
          hour: toNumber(item?.actual_hour ?? item?.plan_hour, 0),
          castName: String(item?.casts?.name || item?.name || ''),
          distanceKm: 0,
          area: normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || ''),
          reason: 'missing_coordinates'
        })),
        selectedAreas: []
      };
      return [];
    }

    const itemsByHour = new Map();
    validPeople.forEach(person => {
      const hour = toNumber(person.actual_hour, 0);
      if (!itemsByHour.has(hour)) itemsByHour.set(hour, []);
      itemsByHour.get(hour).push(person);
    });

    const assignments = [];
    const meta = {
      mode: 'anchor_slot_fixed',
      vehicleCount: safeVehicles.length,
      actualGroupCount: 0,
      keptGroups: [],
      overflowGroups: [],
      evaluations: [],
      totalSeatCapacity: safeVehicles.reduce((sum, v) => sum + Math.max(1, toNumber(v.seat_capacity, 4)), 0),
      totalCastCount: safeItems.length,
      capacityOverflowCount: 0,
      capacityOverflowItems: [],
      selectedAreas: []
    };

    [...itemsByHour.keys()].sort((a, b) => a - b).forEach(hour => {
      const hourPeople = itemsByHour.get(hour) || [];
      const result = buildHourAssignments(origin, safeVehicles, hourPeople);
      meta.actualGroupCount += result.groups.length;
      meta.evaluations.push(...result.evaluations);

      result.groups.forEach(group => {
        const memberIds = (group.members || []).map(person => toNumber(person.id, 0)).filter(Boolean);
        const orderedIds = (group.routeOrder || []).map(person => toNumber(person.id, 0)).filter(Boolean);
        meta.keptGroups.push({
          hour: toNumber(hour, 0),
          groupId: String(group.groupId || ''),
          vehicleId: toNumber(group.vehicleId, 0),
          vehicleName: String(group.vehicleName || ''),
          anchorId: toNumber(group.anchor?.id, 0),
          anchorName: String(group.anchor?.name || ''),
          anchorDistance: Number(group.anchor?.priorityDistance || 0),
          anchorAngle: Number(group.anchor?.angleFromOrigin || 0),
          count: (group.routeOrder || []).length,
          itemIds: orderedIds,
          memberIds,
          routeDistance: Number(group.routeDistance || 0),
          area: String(group.anchor?.destination_area || '')
        });
        meta.selectedAreas.push(`${hour}:${String(group.anchor?.destination_area || '')}:${String(group.anchor?.name || '')}`);
      });

      if (result.overflow.length) {
        meta.overflowGroups.push({
          hour: toNumber(hour, 0),
          group: `overflow_${hour}`,
          count: result.overflow.length,
          itemIds: result.overflow.map(person => toNumber(person.id, 0)).filter(Boolean)
        });
        meta.capacityOverflowCount += result.overflow.length;
        meta.capacityOverflowItems.push(...result.overflow.map(person => ({
          itemId: toNumber(person.id, 0),
          hour: toNumber(hour, 0),
          castName: String(person.name || ''),
          distanceKm: Number(person.priorityDistance || 0),
          area: String(person.destination_area || ''),
          reason: 'no_available_capacity'
        })));
      }

      assignments.push(...buildAssignmentsFromGroups(result.groups, hour));
    });

    if (invalidPeople.length) {
      meta.capacityOverflowCount += invalidPeople.length;
      meta.capacityOverflowItems.push(...invalidPeople.map(item => ({
        itemId: toNumber(item?.id, 0),
        hour: toNumber(item?.actual_hour ?? item?.plan_hour, 0),
        castName: String(item?.casts?.name || item?.name || ''),
        distanceKm: 0,
        area: normalizeAreaLabel(item?.destination_area || item?.cluster_area || item?.planned_area || item?.casts?.area || ''),
        reason: 'missing_coordinates'
      })));
    }

    globalThis.__THEMIS_LAST_OVERFLOW__ = meta;

    return assignments.sort((a, b) => {
      if (toNumber(a.actual_hour, 0) !== toNumber(b.actual_hour, 0)) return toNumber(a.actual_hour, 0) - toNumber(b.actual_hour, 0);
      if (toNumber(a.vehicle_id, 0) !== toNumber(b.vehicle_id, 0)) return toNumber(a.vehicle_id, 0) - toNumber(b.vehicle_id, 0);
      if (toNumber(a.stop_order, 0) !== toNumber(b.stop_order, 0)) return toNumber(a.stop_order, 0) - toNumber(b.stop_order, 0);
      return toNumber(a.item_id, 0) - toNumber(b.item_id, 0);
    });
  }

  return {
    optimizeAssignments,
    pureOptimizeAssignments: optimizeAssignments,
    selectAnchors
  };
})();
