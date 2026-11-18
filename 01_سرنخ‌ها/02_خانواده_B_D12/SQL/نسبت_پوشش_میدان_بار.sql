SELECT m.facility_id, m.name,
       SUM(v.population * sa.coverage_ratio) / (SELECT SUM(population) FROM villages) AS coverage_ratio
FROM facilities m
JOIN facility_service_area sa ON sa.facility_id = m.facility_id
JOIN regions v ON v.id = sa.region_id
WHERE m.facility_type = 'fruit_vegetable_market'
GROUP BY m.facility_id, m.name;