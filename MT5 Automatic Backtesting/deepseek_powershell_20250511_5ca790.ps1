# Add to initialization
SET powerAutomateFlagFile TO "D:\FOREX\pa_control\new_data.flag"  # File to signal Power Automate

# Modify the report saving section to add:
# After successfully saving a report:
CREATE FILE "%powerAutomateFlagFile%"