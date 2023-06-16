#Finds duplicate serials
SELECT tag_number, description, building, area_num, serial_num
FROM assets
WHERE serial_num IN (
    SELECT serial_num
    FROM assets
    GROUP BY serial_num
    HAVING COUNT(serial_num) > 1
)
ORDER BY serial_num