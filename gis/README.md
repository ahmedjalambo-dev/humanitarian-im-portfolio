# GIS layer

`itcn_sites.geojson` (EPSG:4326) holds the 28 sites whose coordinates passed validation. The two sites that failed (one missing GPS, one outside the Gaza Strip) are excluded and logged in the DQ log for re-capture.

Attributes: site_id, site_name, governorate, site_type, households, households_registered, rhu_target, rhu_installed, workers, workdays, avg_progress.

QGIS map: OpenStreetMap basemap, sites graduated by `households`, labelled by `site_id`, exported as `docs/screenshots/qgis_site_map.png`.
