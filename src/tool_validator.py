# =============================================================================
# SOURCE MIRROR - READ ONLY REFERENCE COPY
# -----------------------------------------------------------------------------
# Extracted from: Geoprocessing Tool Grouping Excel Export.atbx
#   archive path: GeoprocessingToolGroupingExcelExportV14.tool/tool.script.validate.py
#
# Published so the code is readable and searchable on GitHub. The .atbx is the
# runnable artifact -- editing this file does NOT change the tool. Edit the
# script inside the toolbox in ArcGIS Pro, then re-run:
#     python tools/extract_atbx_scripts.py
# =============================================================================

class ToolValidator:
  # Class to add custom behavior and properties to the tool and tool parameters.

    def __init__(self):
        # Set self.params for use in other validation methods.
        self.params = arcpy.GetParameterInfo()

    def initializeParameters(self):
        # Customize parameter properties. This method gets called when the
        # tool is opened.
        return

    # Some options need to be hidden/disabled if another specific option is selected.
    def updateParameters(self):
    
        # Parameter indices (for reference)
        # 0 = feature_layer
        # 1 = Export_All_Fields
        # 2 = Fields_To_Export
        # 3 = Group_By_Field
        # 4 = Include_Group_Field_In_Output
        # 5 = Use_Domain_Descriptions
        # 6 = Use_Field_Aliases
        # 7 = Output_XLSX
        # 8 = EXPORT_MODE
        # 9 = Add_TOC
        # 10 = AUTO_WIDTH
    
        export_all_param   = self.params[1]
        field_picker_param = self.params[2]
        export_mode_param  = self.params[8]
        add_toc_param      = self.params[9]
    
        # ------------------------------
        # 1. Disable field picker when Export_All_Fields = True
        # ------------------------------
        if export_all_param.value:
            field_picker_param.enabled = False
        else:
            field_picker_param.enabled = True
    
        # ------------------------------
        # 2. Hide OR show the Add_TOC parameter based on EXPORT_MODE
        # ------------------------------
        if export_mode_param.value == "single_sheet":
            # In single-sheet mode, no Table of Contents sheet is created,
            # so hide the parameter entirely.
            add_toc_param.visible = False
            add_toc_param.enabled = False
            add_toc_param.value   = False  # Optional: auto-clear the checkbox
        else:
            # In multi-sheet mode (“sheets”), show and enable the Add_TOC option.
            add_toc_param.visible = True
            add_toc_param.enabled = True
    
        return



    def updateMessages(self):
        # Modify the messages created by internal validation for each tool
        # parameter. This method is called after internal validation.
        return

    # def isLicensed(self):
    #     # Set whether the tool is licensed to execute.
    #     return True

    # def postExecute(self):
    #     # This method takes place after outputs are processed and
    #     # added to the display.
    #     return
