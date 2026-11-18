SELECT c.facility_id, c.name, r.threshold_value
FROM facilities c
JOIN spatial_validation_rules r ON r.facility_type = 'cemetery'
JOIN water_bodies w ON ST_DWithin(c.geometry::geography, w.geometry::geography, r.threshold_value)
WHERE c.facility_type = 'cemetery' AND r.target_layer = 'water_surface';