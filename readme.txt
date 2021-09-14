v1.0.1 release note
  - first release to match manufacturer_item_number of 2 CSV files
  - Random color for every matched item.
  - 2 tabs to include import tool as well.

v1.0.2 release note
  - add export function.
  - insert item_number when manufacturer_item_number matched.
  - skip "none", space, "N/A" fields in manufacturer_item_number.
  - highlight vendor_item_number, item_name is different when manufacturer_item_number matched.

v1.0.3 release note
  - Add check for CSV1 duplicated manufacturer_item_number but different item_name, item_number, vendor_item_number also mark red color
  - fix count bug, only count duplicated items for only one time.

v1.0.4 release note
  - Add export file to support Excel .xlsx file
  - Fix export csv file bug, last column will be missed.

v.1.0.5 release note
  - Fix column text includ character '"', in csv format it should saved as '""'.

v1.0.6 release note
  - Add check for four mandatory fields. - item_number, item_name, vendor_item_number, manufacturer_item_number, must exists in both csv files.
