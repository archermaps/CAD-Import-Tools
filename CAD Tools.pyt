"""
CAD Tools.pyt
Esri - Database Services
Brooke Reams, breams@esri.com
Feb. 22, 2016
- Imports AIA CAD files into the LGIM data model.

Updates:
03/16/2016:     Updated script to read configuration values from config file.
03/17/2016:     Updated script to work with config file with lists using ast module.
                Added roomiden_list to config in order to dynamically update where
                clause for anno to points operation.
06/16/2016:     Updated functionality of merge floor option: if option is checked,
                the building features are created from dissolving the building floor
                features on the BUILDINGKEY field.  If the option is unchecked, the
                building features are created from the layers specified in the config
                under the Building field.

"""

import arcpy, os, sys, traceback, __future__##, ConfigParser, ast, 
##from openpyxl import load_workbook

##class ex(Exception):
##    def __init__(self, value):
##        self.parameter = value
##    def __str__(self):
##        return repr(self.parameter)


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Toolbox"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [CAD2LGIM]


class CAD2LGIM(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "AIA CAD To LGIM"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        param0 = arcpy.Parameter(
        displayName='Input CAD Files',
        name='input_cad',
        datatype='GPValueTable',
        parameterType='Required',
        direction='Input')        
        # Create value table columns and column filters
        param0.columns = [['DEFile', 'CAD File'], ['GPString', 'Buildling ID'], ['GPString', 'Floor Number'], ['GPString', 'Floor ID'],
                          ['GPString', 'Floor Count'], ['GPString', 'Floor Description'], ['GPString', 'Base Elevation'],
                          ['GPString', 'Ceiling Height'], ['GPString', 'Vertical Order']]

        param1 = arcpy.Parameter(
        displayName='LGIM Data Model',
        name='lgim_data_model',
        datatype='DEWorkspace',
        parameterType='Required',
        direction='Input')
        # Filter workspace type to file gdb
        param1.filter.list = ["Local Database"]

        param2 = arcpy.Parameter(
        displayName='Reference Scale',
        name='ref_scale',
        datatype='GPDouble',
        parameterType='Required',
        direction='Input')
        # Default reference scale
        param2.value = 1000

        param3 = arcpy.Parameter(
        displayName='Spatial Reference',
        name='spatial_ref',
        datatype='GPSpatialReference',
        parameterType='Required',
        direction='Input')
##        # Default spatial reference
##        param3.value = 26911

        param4 = arcpy.Parameter(
        displayName='Output QC Geodatabase',
        name='output_qc_gdb',
        datatype='DEWorkspace',
        parameterType='Required',
        direction='Input')
        # Filter workspace type to file gdb
        param4.filter.list = ["Local Database"]

        param5 = arcpy.Parameter(
        displayName='Building Interior Spaces Import Polgyons',
        name='bldgintspace_import_polys',
        datatype='GPBoolean',
        parameterType='Optional',
        direction='Input')
        param5.value = True
        
        param6 = arcpy.Parameter(
        displayName='Minimum Area',
        name='min_area',
        datatype='GPDouble',
        parameterType='Optional',
        direction='Input')

        param7 = arcpy.Parameter(
        displayName='Merge Floor to Building Footrpint',
        name='merge_floor',
        datatype='GPBoolean',
        parameterType='Optional',
        direction='Input')

        params = [param0, param1, param2, param3, param4, param5, param6, param7]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

        if not parameters[0].hasBeenValidated and not parameters[0].altered:
            prop_config = os.path.join(sys.path[0], "Properties.csv")

            # Check if config exists
            if os.path.exists(prop_config):
                value_tbl = []
                with open(prop_config, "r") as f_in:
                    # Read first line
                    f_in.readline()
                    # Store layer names in lists
                    for row in f_in:
                        line = row.split(",")
                        values = []
                        values.append(line[0])
                        values.append(line[1])
                        values.append(line[2])
                        values.append(line[3])
                        values.append(line[4])
                        values.append(line[5])
                        values.append(line[6])
                        values.append(line[7])
                        values.append(line[8].replace("\n", ""))

                        value_tbl.append(values)

                    parameters[0].values = value_tbl
            
##            xls_config = os.path.join(sys.path[0], "config.xlsx")
##            if os.path.exists(xls_config):
##                value_tbl = []
##
##                # Open  xlsx and retrieve CAD file data
##                wb = load_workbook(xls_config, data_only=True)
##                if "Properties" in wb.get_sheet_names():
##                    sheet = wb.get_sheet_by_name("Properties")
##                    iter = sheet.iter_rows()
##                    # Skip first row (contains field names)
##                    iter.next()
##                    for row in iter:
##                        values = []
##                        values.append(row[0].value)
##                        values.append(row[1].value)
##                        values.append(row[2].value)
##                        values.append(row[3].value)
##                        values.append(row[4].value)
##                        values.append(row[5].value)
##                        values.append(row[6].value)
##                        values.append(row[7].value)
##                        values.append(row[8].value)
##                        
##                        value_tbl.append(values)
##
##                parameters[0].values = value_tbl
        
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def getFieldMap(self, input_fc, input_fld, output_fld):
        # Create field map object
        fm = arcpy.FieldMap()

        # Add input field to field map
        fm.addInputField(input_fc, output_fld)

        # Set output field properties
        output_fld = fm.outputField
        output_fld.name = input_fld
        fm.outputField = output_fld

        return fm


    def addCommonFields(self, fc, common_flds):
        for fld, info in common_flds.items():
            fld_found = arcpy.ListFields(fc, fld)
            if not fld_found:
                arcpy.AddField_management(fc, fld, info["type"])
            if info["type"].lower() == "text":
                try:
                    arcpy.CalculateField_management(fc, fld, "'" + info["value"] + "'", "PYTHON_9.3")
                except:
                    arcpy.CalculateField_management(fc, fld, "None", "PYTHON_9.3")
                    pass
            else:
                try:
                    arcpy.CalculateField_management(fc, fld, str(info["value"]), "PYTHON_9.3")
                except:
                    arcpy.CalculateField_management(fc, fld, "None", "PYTHON_9.3")
                    pass

        return


    def addBuildingInteriorSpaceFields(self, fc, lyr_count, product):
        # Add/calc SpaceType and SpaceID
        for fld in ["SPACEID", "SPACETYPE"]:
            fld_found = arcpy.ListFields(fc, fld)
            if not fld_found:
                arcpy.AddField_management(fc, fld, "TEXT")
            space_lyr = "space_lyr"  + str(lyr_count)
            if fld.lower() == "spacetype":
                arcpy.MakeFeatureLayer_management(fc, space_lyr, "Layer = 'A-AREA-TYPE'")
                if product == "Desktop":
                    arcpy.CalculateField_management(space_lyr, fld, "!TextString!", "PYTHON_9.3")
                else:
                    arcpy.CalculateField_management(space_lyr, fld, "!Text!", "PYTHON_9.3")
            else:
                arcpy.MakeFeatureLayer_management(fc, space_lyr, "Layer = 'A-AREA-IDEN'")
                if product == "Desktop":
                    arcpy.CalculateField_management(space_lyr, fld, "!TextString!", "PYTHON_9.3")
                else:
                    arcpy.CalculateField_management(space_lyr, fld, "!Text!", "PYTHON_9.3")

        return


    def getBuildingOrientation(self, working_fc, lyr, target_fc):
        geom_fc = "in_memory\\min_bounding_geom"
        arcpy.MinimumBoundingGeometry_management(working_fc, geom_fc, "CONVEX_HULL", mbg_fields_option="MBG_FIELDS")
        # Add Rotation field to layer
        fld_found = arcpy.ListFields(working_fc, "Rotation")
        if not fld_found:
            arcpy.AddField_management(working_fc, "Rotation", "DOUBLE")
        # Join minimum bounding geometry to layer to calc orientation field
        oid_fld = arcpy.Describe(working_fc).OIDFieldName
        arcpy.AddJoin_management(lyr, oid_fld, geom_fc, "ORIG_FID_1")
        # Calc Orientation value
        arcpy.CalculateField_management(lyr, "Rotation", "!min_bounding_geom.MBG_Orientation!", "PYTHON_9.3")
        # Remove join
        arcpy.RemoveJoin_management(lyr, "min_bounding_geom")
        # If Buildling feature class does not have Rotation field, add it
        fld_found = arcpy.ListFields(target_fc, "Rotation")
        if not fld_found:
            arcpy.AddField_management(target_fc, "Rotation", "DOUBLE")

        return


    def qcFindIdenticalFeatures(self, fc, output_qc_gdb, fds_name, lyr_count):
        # Find identical features
        iden_tbl = "in_memory//identicals"
        arcpy.FindIdentical_management(fc, iden_tbl, "Shape;Layer", output_record_option="ONLY_DUPLICATES")

        # If identical features were found, determine which features will be deleted by Delete Identical tool
        count = int(arcpy.GetCount_management(iden_tbl).getOutput(0))
        if count > 0:
            id_list = []
            with arcpy.da.SearchCursor(iden_tbl, ["FEAT_SEQ"], sql_clause=(None, "GROUP BY FEAT_SEQ")) as cur:
                for row in cur:
                    id_list.append(row[0])

            d = {}
            for id in id_list:
                where = "FEAT_SEQ = " + str(id)
                l = []
                with arcpy.da.SearchCursor(iden_tbl, ["IN_FID"], where) as cur:
                    for row in cur:
                        l.append(row[0])
                        d[id] = l

            for k, v in d.items():
                v.pop(0)

            feat_list = []
            for i in d.values():
                for j in i:
                    feat_list.append(j)

            if len(feat_list) > 0:
                # Make feature layer from identical features found
                where = "OBJECTID IN " + str(feat_list).replace("[", "(").replace("]", ")")
                iden_lyr = "iden_feat"  + str(lyr_count)
                arcpy.MakeFeatureLayer_management(fc, iden_lyr, where)
                # Copy features to qc fc
                identical_qc_fc = os.path.join(output_qc_gdb, os.path.basename(fc) + "_" + fds_name + "_Identical_QC")
                arcpy.CopyFeatures_management(iden_lyr, identical_qc_fc)
                # Delete identical features
                arcpy.DeleteFeatures_management(iden_lyr)

        return


    def qcFindNullGeometry(self, fc, output_qc_gdb, fds_name, lyr_count):
        # Check geometry
        geom_tbl = "in_memory\\geom"
        arcpy.CheckGeometry_management(fc, geom_tbl)

        # If bad geometries were found, determine which features have null geometries
        # Features with null geometries will be deleted by the Repair Geometry tool, so
        # want to copy them to a qc fc before they are deleted
        count = int(arcpy.GetCount_management(geom_tbl).getOutput(0))
        if count > 0:
            null_list = []
            with arcpy.da.SearchCursor(geom_tbl, ["FEATURE_ID"], "PROBLEM = 'null geometry'") as cur:
                for row in cur:
                    null_list.append(row[0])

            if len(null_list) > 0:
                # Make feature layer from null geometries found
                where = "OBJECTID IN " + str(null_list).replace("[", "(").replace("]", ")")
                null_lyr = "null_geom"  + str(lyr_count)
                arcpy.MakeFeatureLayer_management(fc, null_lyr, where)
                # Copy features to qc fc
                null_qc_fc = os.path.join(output_qc_gdb, os.path.basename(fc) + "_" + fds_name + "_NullGeom_QC")
                arcpy.CopyFeatures_management(null_lyr, null_qc_fc)

        return


    def qcAnno(self, polygon_fc, anno_pt_fc, temp_gdb, output_qc_gdb, fds_name, lyr_count):
        # Get path to output fc
        anno_qc_fc = os.path.join(output_qc_gdb, fds_name + "_Anno_QC")
        # Make sub layer of anno where "Layer = 'A-AREA-IDEN'"
        anno_lyr = "anno_lyr" + str(lyr_count)
        arcpy.MakeFeatureLayer_management(anno_pt_fc, anno_lyr, "Layer = 'A-AREA-IDEN'")
        # Find building interior spaces that have more than one anno point or no anno point
        output_join = os.path.join(temp_gdb, "CAD_sj")
        arcpy.SpatialJoin_analysis(polygon_fc, anno_lyr, output_join, match_option="CONTAINS")
        # Export features where Join Count != 1
        join_lyr = "join_lyr"  + str(lyr_count)
        arcpy.MakeFeatureLayer_management(output_join, join_lyr, "JOIN_COUNT <> 1")
        count = arcpy.GetCount_management(join_lyr).getOutput(0)
        if int(count) > 0:
            arcpy.CopyFeatures_management(join_lyr, anno_qc_fc)

        return



    def qcMinArea(self, min_area, output_qc_gdb, fds_name):
        # Get path to output fc
        min_qc_fc = os.path.join(output_qc_gdb, fds_name + "_MinArea_QC")
        if not arcpy.Exists(min_qc_fc):
            arcpy.CopyFeatures_management(min_area, min_qc_fc)
        else:
            arcpy.Append_management(min_area, min_qc_fc, "NO_TEST")

        return


    def qcSlivers(self, slivers, output_qc_gdb, fds_name):
        # Get path to output fc
        slivers_qc_fc = os.path.join(output_qc_gdb, fds_name + "_Slivers_QC")
        if not arcpy.Exists(slivers_qc_fc):
            arcpy.CopyFeatures_management(slivers, slivers_qc_fc)
        else:
            arcpy.Append_management(slivers, slivers_qc_fc, "NO_TEST")

        return


    def execute(self, parameters, messages):
        """The source code of the tool."""
        val_tbl = parameters[0].value
        target_gdb = parameters[1].valueAsText
        ref_scale = parameters[2].value
        sr = parameters[3].value
        output_qc_gdb = parameters[4].valueAsText
        bldgintspace_import_polys = parameters[5].valueAsText
        min_area = parameters[6].value
        merge_floor = parameters[7].valueAsText

        # Overwrite existing output
        arcpy.env.overwriteOutput = 1

        # Config file
        layers_config = os.path.join(sys.path[0], "Layers.csv")##os.path.join(sys.path[0], "config.xlsx")

        # Initialize layers lists
        tgtbldg_list = []
        tgtbldgflr_list = []
        tgtbldginterior_list = []
        tgtbldgfloorplan_list = []
        roomiden_list = []

        # Check if config exists
        if os.path.exists(layers_config):
            with open(layers_config, "r") as f_in:
                # Read first line
                f_in.readline()
                # Store layer names in lists
                for row in f_in:
                    line = row.split(",")
                    if line[0]:
                        tgtbldg_list.append(str(line[0]))
                    if line[1]:
                        tgtbldgflr_list.append(str(line[1]))
                    if line[2]:
                        tgtbldgfloorplan_list.append(str(line[2]))
                    if line[3]:
                        tgtbldginterior_list.append(str(line[3]))
                    if line[4]:
                        roomiden_list.append(str(line[4]).replace("\n", ""))
                        
##            # Open  xlsx and retrieve layer data
##            wb = load_workbook(xls_config, data_only=True)
##            sheet = wb.get_sheet_by_name("Layers")
##            iter = sheet.iter_rows()
##            # Skip first row (contains field names)
##            iter.next()
##            for row in iter:
##                if row[0].value:
##                   tgtbldg_list.append(str(row[0].value))
##                if row[1].value:
##                    tgtbldgflr_list.append(str(row[1].value))
##                if row[2].value:
##                    tgtbldgfloorplan_list.append(str(row[2].value))
##                if row[3].value:
##                    tgtbldginterior_list.append(str(row[3].value))
##                if row[4].value:
##                    roomiden_list.append(str(row[4].value))


        # Error out - config not found
        else:
            arcpy.AddError("Cannot find configuration file: Properties.csv.  Ensure configuration file is stored in same folder as script.")
            sys.exit()

        try:
            # Check if product is Desktop or Prop - difference in field name in Anno layer (no field called TextString when anno layer is created in Pro)
            product = arcpy.GetInstallInfo()["ProductName"]
            
            # Initialize temp gdb
            temp_gdb = None

            # Define output feature classes in target gdb
            target_bldg = os.path.join(target_gdb, "Building")
            target_bldgfloor = os.path.join(target_gdb, "BuildingFloor")
            target_bldgfloorplan = os.path.join(target_gdb, "BuildingFloorPlanLine")
            target_bldgfloorplanpub = os.path.join(target_gdb, "BuildingFloorPlanPublish")
            target_bldginterior = os.path.join(target_gdb, "BuildingInteriorSpace")

            # Define layers to export by target feature class
            export_dict = {
                target_bldgfloor: tgtbldgflr_list,
                target_bldgfloorplan: tgtbldgfloorplan_list,
                target_bldginterior: tgtbldginterior_list}
            # Only include Building fc if merge option is false
            if merge_floor != "true":
                export_dict[target_bldg] = tgtbldg_list

            # Define mappings from user input to LGIM fieldss
            input_mappings = {target_bldg: {"BUILDINGID": "BUILDINGID", "FLOORCOUNT": "FLOORCOUNT", "ROTATION": "ROTATION"},
                              target_bldgfloor: {"BUILDINGKEY": "BUILDINGID", "FLOORID": "FLOORID", "DESCRIP": "DESCRIP", "BASEELEV": "BASEELEV",
                                                 "VERTORDER": "VERTORDER", "FLOOR": "FLOORNUM"},
                              target_bldgfloorplan: {"BUILDINGKEY": "BUILDINGID", "FLOORID": "FLOORID", "BASEELEV": "BASEELEV", "HEIGHT": "HEIGHT",
                                                     "FLOOR": "FLOORNUM", "LINETYPE": "LAYER"},
                              target_bldgfloorplanpub: {"BUILDINGKEY": "BUILDINGID", "FLOORID": "FLOORID", "BASEELEV": "BASEELEV", "HEIGHT": "HEIGHT",
                                                        "FLOOR": "FLOORNUM"},
                              target_bldginterior: {"FLOOR": "FLOORNUM", "BUILDING": "BUILDINGID", "FLOORKEY": "FLOORID", "DESCRIP": "DESCRIP",
                                                    "SPACEID": "SPACEID", "SPACETYPE": "SPACETYPE", "BASEELEV": "BASEELEV", "CEILINGHEIGHT": "HEIGHT",
                                                    "FLOORKEY": "FLOORID"}}


            # Loop through CAD files in value table
            lyr_count = 0  # ArcGIS Pro bug with CAD2GDB tool prevents layers from being overwritten - numbering layers for each iteration of the tool to work around bug
            for value in val_tbl:
                lyr_count += 1
                input_cad = str(value[0])
                if str(value[1]) != "":
                    building_id = str(value[1])
                else:
                    building_id = None
                if value[2]:
                    floor_num = str(value[2])
                else:
                    floor_num = None
                if value[3]:
                    floor_id = str(value[3])
                else:
                    floor_id = None
                if value[4]:
                    floor_count = str(value[4])
                else:
                    floor_count = None
                if value[5]:
                    desc = str(value[5])
                else:
                    desc = None
                if str(value[6]) != "":
                    base_elev = float(value[6])
                else:
                    base_elev = None
                if str(value[7]) != "":
                    ceiling_height = float(value[7])
                else:
                    ceiling_height = None
                if str(value[8]) != "":
                    vertical_order = int(value[8])
                else:
                    vertical_order = None
                    
                print(os.path.basename(input_cad) + ":")
                arcpy.AddMessage(os.path.basename(input_cad) + ":")
                
                common_flds = {"BUILDINGID": {"type": "TEXT", "value": building_id}, "FLOORNUM": {"type": "TEXT", "value": floor_num},
                               "FLOORID": {"type": "TEXT", "value": floor_id}, "DESCRIP": {"type": "TEXT", "value": desc},
                               "FLOORCOUNT": {"type": "TEXT", "value": floor_count},"BASEELEV": {"type": "DOUBLE", "value": base_elev},
                               "HEIGHT": {"type": "DOUBLE", "value": ceiling_height}, "VERTORDER": {"type": "LONG", "value": vertical_order}}
                    
                # Create file gdb if doesn't already exist
                temp_gdb = os.path.join(arcpy.env.scratchFolder, "TempCAD.gdb")
                if not arcpy.Exists(temp_gdb):
                    arcpy.CreateFileGDB_management(arcpy.env.scratchFolder, "TempCAD.gdb")
                # Import CAD to temp gdb
                file_name = str(os.path.splitext(os.path.basename(input_cad))[0])
                fds_name = "CAD_" + "".join(i for i in file_name if i.isalnum())
                arcpy.CADToGeodatabase_conversion(input_cad, temp_gdb, fds_name, ref_scale, sr)

                # Get paths to fcs in temp gdb
                temp_polygon = os.path.join(temp_gdb, fds_name, "Polygon")
                temp_anno = os.path.join(temp_gdb, fds_name, "Annotation")
                temp_polyline = os.path.join(temp_gdb, fds_name, "Polyline")

                # Create points from annotation
                anno_lyr = "anno_lyr" + str(lyr_count)
                arcpy.MakeFeatureLayer_management(temp_anno, anno_lyr, "Layer IN " + str(tuple(roomiden_list)))
                anno_to_point = os.path.join(temp_gdb, "Anno2Point")
                arcpy.FeatureToPoint_management(anno_lyr, anno_to_point)

                # Add/calc common fields to temp CAD fc
                self.addCommonFields(temp_polyline, common_flds)
                self.addCommonFields(temp_polygon, common_flds)

                # Clean up any issues with input CAD features
                for temp_fc in [temp_polyline, temp_anno]:
                    # QC CHECK - Find null geometry
                    self.qcFindNullGeometry(temp_fc, output_qc_gdb, fds_name, lyr_count)
                    # Repair bad geometry
                    arcpy.RepairGeometry_management(temp_fc, "DELETE_NULL")
                    # QC CHECK - Find Identical features
                    self.qcFindIdenticalFeatures(temp_fc, output_qc_gdb, fds_name, lyr_count)
                    

                # Loop through dictionary of target fcs and layers
                for target_fc, layer_types in export_dict.items():
##                    for layer in layer_types:
                    # Reset field mappings object
                    fms = None
                    # Flag for found features
                    feat_found = False
                    # Flag for BuildingInteriorSpace polygons
                    poly_found = False
                    # First check if layer is in polygon - if not, then use features from polyline
                    if os.path.basename(target_fc).lower() == "buildinginteriorspace" and bldgintspace_import_polys == "true":
                        # Create feature layer of CAD layer in polygon
##                        arcpy.MakeFeatureLayer_management(temp_polygon, "lyr", "Layer = '{}'".format(layer))
                        layers = str(tuple(layer_types))
                        if len(layer_types) == 1:
                            layers = layers.replace(",", "")
                        lyr = "lyr" + str(lyr_count)
                        arcpy.MakeFeatureLayer_management(temp_polygon, lyr, "Layer IN " + layers)
                        count = arcpy.GetCount_management(lyr).getOutput(0)
                        if int(count) > 0:
                            feat_to_polygon = os.path.join(temp_gdb, "Feature2Polygon")
                            arcpy.FeatureToPolygon_management(lyr, feat_to_polygon, label_features=anno_to_point)
                            # Add/calc common fields to new output polygon fc
                            self.addCommonFields(feat_to_polygon, common_flds)
                            lyr = "lyr" + str(lyr_count)
                            arcpy.MakeFeatureLayer_management(feat_to_polygon, lyr)
                            feat_found = True
                            working_fc = feat_to_polygon
                            poly_found = True
                    if not poly_found:
                        # Create feature layer
##                        arcpy.MakeFeatureLayer_management(temp_polyline, "lyr", "Layer = '{}'".format(layer))
                        layers = str(tuple(layer_types))
                        if len(layer_types) == 1:
                            layers = layers.replace(",", "")
                        lyr = "lyr" + str(lyr_count)
                        arcpy.MakeFeatureLayer_management(temp_polyline, lyr, "Layer IN" + layers)
                        # Get count
                        count = arcpy.GetCount_management(lyr).getOutput(0)
                        if int(count) > 0:
                            feat_found = True
                            # If target is polygon, need to convert CAD features to polygons
                            if os.path.basename(target_fc).lower() != "buildingfloorplanline":  # Polygon
                                polyline_to_polygon = os.path.join(temp_gdb, "Polyline2Polygon")
                                arcpy.FeatureToPolygon_management(lyr, polyline_to_polygon, label_features=anno_to_point)
                                # Add/calc common fields to new output polygon fc
                                self.addCommonFields(polyline_to_polygon, common_flds)
                                arcpy.MakeFeatureLayer_management(polyline_to_polygon, lyr)
                                working_fc = polyline_to_polygon

                            else:
                                # Dissolve lines on floor, buildingkey, and linetype
                                dissolve_fc = os.path.join(temp_gdb, "bldgfloorplanline_dissolve")
                                arcpy.Dissolve_management(lyr, dissolve_fc, "Linetype;Layer;BUILDINGID;FLOORID", multi_part="SINGLE_PART")
                                # Add/calc common fields to new output polygon fc
                                self.addCommonFields(dissolve_fc, common_flds)
                                arcpy.MakeFeatureLayer_management(dissolve_fc, lyr)
                                working_fc = dissolve_fc

                    if feat_found:
                        if os.path.basename(target_fc).lower() == "buildinginteriorspace":
                            # If target fc is BuildingInteriorSpace then delete features less than minimum area
                            if min_area:
                                # Make feature layer
                                where = "SHAPE_AREA < " + str(min_area)
                                min_area_lyr = "min_area"  + str(lyr_count)
                                arcpy.MakeFeatureLayer_management(working_fc, min_area_lyr, where)
                                # Export deleted features to qc fc
                                self.qcMinArea(min_area_lyr, output_qc_gdb, fds_name)
                                # Delete the features from the temp fc so they are not appended to the BuildingInteriorSpace fc
                                arcpy.DeleteFeatures_management(min_area_lyr)

                            # Delete slivers where SHAPE_Length >= (SHAPE_Area*2)
                            where = "SHAPE_LENGTH > (SHAPE_AREA * 2)"
                            slivers_lyr = "slivers"  + str(lyr_count)
                            arcpy.MakeFeatureLayer_management(working_fc, slivers_lyr, where)
                            self.qcSlivers(slivers_lyr, output_qc_gdb, fds_name)
                            arcpy.DeleteFeatures_management(slivers_lyr)
                            
                            # Add and calc fields for CAD attributes
                            self.addBuildingInteriorSpaceFields(working_fc, lyr_count, product)
                            # Report on duplicate/missing SpaceIDs
                            self.qcAnno(working_fc, anno_to_point, temp_gdb, output_qc_gdb, fds_name, lyr_count)
                        # If target fc is Building, get orientation from Minimum Bounding Geometry tool
                        elif os.path.basename(target_fc).lower() == "building":
                            self.getBuildingOrientation(working_fc, lyr, target_fc)
                        # Get field mappings from user input
                        if target_fc in input_mappings.keys():
                            # Create field mappings object
                            fms = arcpy.FieldMappings()
                            fld_dict = input_mappings[target_fc]
                            for input_fld, output_fld in fld_dict.items():
                                fm = self.getFieldMap(working_fc, input_fld, output_fld)
                                # Add field map to field mappings
                                fms.addFieldMap(fm)

                        print("\tAppending layers to " + os.path.basename(target_fc))
                        arcpy.AddMessage("\tAppending layers to " + os.path.basename(target_fc))
                        # Append data to target fc
                        arcpy.Append_management(lyr, target_fc, "NO_TEST", fms)
                        # If target fc is BuildingFloorPlanLine, append data to BuildlingFloorPlanPublish
                        if os.path.basename(target_fc).lower() == "buildingfloorplanline":
                            fms_floorplanpub = arcpy.FieldMappings()
                            fld_dict = input_mappings[target_bldgfloorplanpub]
                            for input_fld, output_fld in fld_dict.items():
                                fm = self.getFieldMap(working_fc, input_fld, output_fld)
                                fms_floorplanpub.addFieldMap(fm)
                            print("\tAppending layers to " + os.path.basename(target_bldgfloorplanpub))
                            arcpy.AddMessage("\tAppending layers to " + os.path.basename(target_bldgfloorplanpub))
                            arcpy.Append_management(lyr,  target_bldgfloorplanpub, "NO_TEST", fms_floorplanpub)
                    else:
                        print("\tNo features to append to " + os.path.basename(target_fc))
                        arcpy.AddMessage("\tNo features to append to " + os.path.basename(target_fc))

                # Delete everything in temp gdb
                # Would normally just overwrite existing gdb next time script is run,
                #    but CAD to GDB GP tool puts some sort of lock on the gdb that
                #    is not removed until ArcMap is closed.
                arcpy.env.workspace = temp_gdb

                for dataset in arcpy.ListDatasets():
                    arcpy.Delete_management(dataset)

                for fc in arcpy.ListFeatureClasses():
                    arcpy.Delete_management(fc)

            # If merge option is true, dissolve BuildingFloor layers and append dissolved feature in Building fc
            if merge_floor == "true":
                dissovle_fc = "in_memory\\dissolve"
                arcpy.Dissolve_management(target_bldgfloor, dissolve_fc, ["BUILDINGKEY"])
##                    # Get building orientation
##                    self.getBuildingOrientation(working_fc, lyr, target_fc)
                # Add/calc common fields to merge fc
                self.addCommonFields(dissolve_fc, common_flds)
                # Append dissolved feature to Building fc
                print("Appending layers to " + os.path.basename(target_bldg))
                arcpy.AddMessage("Appending layers to " + os.path.basename(target_bldg))
                arcpy.Append_management(dissolve_fc, target_bldg, "NO_TEST")


##        except ex, (instance):
##            print(instance.parameter)
##            arcpy.AddError(instance.parameter)

        except arcpy.ExecuteError:
            # Get the geoprocessing error messages
            msgs = arcpy.GetMessage(0)
            msgs += arcpy.GetMessages(2)

            # Write gp error messages to log
            print(msgs + "\n")
            arcpy.AddError(msgs + "\n")


        except:
            # Get the traceback object
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]

            # Concatenate information together concerning the error into a message string
            if product == "Desktop":
                pymsg = tbinfo + "\n" + str(sys.exc_type)+ ": " + str(sys.exc_value)
            else:
                pymsg = tbinfo + "\n" + str(sys.exc_info()[1])

            # Write Python error messages to log
            print(pymsg + "\n")
            arcpy.AddError(pymsg + "\n")

        finally:
            # Delete everything in temp gdb
            # Would normally just overwrite existing gdb next time script is run,
            #    but CAD to GDB GP tool puts some sort of lock on the gdb that
            #    is not removed until ArcMap is closed.
            arcpy.env.workspace = temp_gdb

            for dataset in arcpy.ListDatasets():
                arcpy.Delete_management(dataset)

            for fc in arcpy.ListFeatureClasses():
                arcpy.Delete_management(fc)


        return
