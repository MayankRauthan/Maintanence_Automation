import automation_vishal as av
import automation_gurmeet as ag

# put the source file name, target file name in the below function call
# Put in regard to gurmeet's report
output_path =ag.orchestration("Maintenance-Data_17 July.xlsx", "Haleon -Network planned maintenance tracker Week 31_29 July_2026.xlsx", "Haleon WAN Inventory.xlsx")

# put target file name in below function call;
# Put in regard to vishal's report
av.orchestration(output_path, "OBS - Carrier Planned Maintenance Tracker -  July_2026.xlsx" ,  "Haleon WAN Inventory.xlsx")