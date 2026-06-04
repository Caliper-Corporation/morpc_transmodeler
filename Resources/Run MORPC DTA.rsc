/*
The environment and relative pathing assumes that the two separate repos
are in separate folders. For example, if the root project folder is: 

C:\projects\MORPC

Then there would be two subfolders:

C10 (CTRAMP folder)
DTA (TransModeler folder)

This script would be in DTA\Resources\Run MORPC DTA.rsc

The C10 scenario folder would be in

C10\ABM\SCEN\Base\MOR

*/

Macro "run_dta"
    
    // Collect arguments and create folder variables
    // outer_iteration = GetEnvironmentVariable("ITER")
    scenario_name = GetEnvironmentVariable("SCEN_NAME")
    if scenario_name = null then scenario_name = "Base"
    region = GetEnvironmentVariable("REGION")
    if region = null then region = "MOR"
    inner = GetEnvironmentVariable("INNER")
    outer = GetEnvironmentVariable("OUTER")
    first_iteration = if inner = "1" and outer = "1" then "True" else "False"
    ui = GetInterface()
    {dir, path, file, ext} = SplitPath(ui)
    model_dir = dir + path + "..\\.." // root project folder
    scen_dir = model_dir + "\\C10\\ABM\\SCEN\\" + scenario_name + "\\" + region

    // Convert the trip list into the format required by TransModeler
    // for morpc, only keep taxi trips from this file
    in_csv = scen_dir + "\\CT-RAMP\\Core\\TripList.csv"
    out_csv = model_dir + "\\DTA\\Demand\\TripList.csv"
    mode_filter = "Select * where (mode = 16 and jointTripRole <= 1)"
    RunMacro("convert trip list", in_csv, out_csv, mode_filter, 10000000)

    // Convert the cartracker file
    in_csv = scen_dir + "\\CT-RAMP\\carTrack\\disaggregateCarUse_carTracker.csv"
    out_csv = model_dir + "\\DTA\\Demand\\disaggregateCarUse_carTracker.csv"
    RunMacro("convert trip list", in_csv, out_csv, , 1)

    // Convert DCOM list
    in_csv = scen_dir + "\\CT-RAMP\\dcom\\DCOMVehicleTrips.csv"
    out_csv = model_dir + "\\DTA\\Demand\\DCOMVehicleTrips.csv"
    RunMacro("convert DCOM trip list", in_csv, out_csv )

    // Run the DTA
    rm = CreateObject("TSM.RunManager")
    smp = model_dir + "\\DTA\\MORPC_DTA.smp"
    rm.OpenSimulationProject(smp, )
    rm.SuppressAllWarnings()
    rm.SetSimulationRunMode("Dynamic Traffic Assignment")
    if !first_iteration then do
        if !rm.LoadPathTables("Demand") then do
            ShowMessage("Failed to load path tables from first iteration.")
            goto quit
        end
    end
    rm.RunSimulation()  // this will start the DTA

    // Create dynamic skims
    RunMacro("Create Dynamic Skims", rm)

    // // Copy the trajectory csv into the appropriate directory
    // trip_file = rm.GetTripTable()
    // {dir, path, name, ext} = SplitPath(trip_file)
    // in_csv = dir + path + "trajMiningInfo.csv"
    // out_dir = model_dir + "\\..\\c10\\outer" + outer_iteration + 
    //     "\\abmData"
    // if GetDirectoryInfo(out_dir, "All") = null then CreateDirectory(out_dir)
    // out_csv = out_dir + "\\trajMiningInfo.csv"
    // CopyFile(in_csv, out_csv)

quit:
    rm.CloseSimulationProject()
    Exit()
endmacro

/* 
Format of CT-RAMP table:
mode:
1 SOV
2 HOV2/driver
3 HOV3+/driver
4 HOV/passenger (not assigned)
5-10 Transit modes (not assigned)
11 Walk (not assigned)
12 Bike (not assigned)
13 Taxi (assigned if driver trip)
14 School bus (not assigned)

jointTripRole:
0 not a joint trip
1 driver of joint trip
2 passenger of joint trip
 */

Macro "test convert trip list taz"
    RunMacro("test convert trip list", "TAZ")
endmacro

Macro "test convert trip list" (od_geography)
    on escape do return() end
    in_csv = ChooseFile({{"Trip List", "*.csv"}, {"Trip List", "*.bin"}}, "Choose a trip list file", )
    on escape default
    parts = SplitPath(in_csv)
    out_folder = ChooseDirectory("Choose the folder in which to save the converted trip list.", {{"Initial Directory", parts[1] + parts[2]}})
    out_csv = out_folder + "\\" + parts[3] + parts[4]
    // for morpc, only keep taxi trips from this file
    mode_filter = "Select * where (mode = 16 and jointTripRole <= 1)"
    RunMacro("convert trip list", in_csv, out_csv, mode_filter, 10000000, od_geography)
    ShowMessage("done")
endmacro

Macro "test convert cartracker taz"
    RunMacro("test convert cartracker", "TAZ")
endmacro

Macro "test convert cartracker" (od_geography)
    on escape do return() end
    in_csv = ChooseFile({{"Trip List", "*.csv"}, {"Trip List", "*.bin"}}, "Choose a trip list file", )
    on escape default
    parts = SplitPath(in_csv)
    out_folder = ChooseDirectory("Choose the folder in which to save the converted trip list.", {{"Initial Directory", parts[1] + parts[2]}})
    out_csv = out_folder + "\\" + parts[3] + parts[4]
    RunMacro("convert trip list", in_csv, out_csv, , 1, od_geography)
    ShowMessage("done")
endmacro

Macro "convert trip list" (in_csv, out_csv, mode_filter, start_id, od_geography)

    // Use MAZs or TAZs?
    if od_geography = "TAZ" then
        {oriField, desField} = {"origTaz", "destTaz"}
    else
        {oriField, desField} = {"origMaz", "destMaz"}

    b_car_tracker = Lower(in_csv) contains "cartracker"

    if in_csv = null then Throw("'in_csv' not provided")
    if out_csv = null then Throw("'out_csv' not provided")

    // Open the trip list.
    vw_csv = OpenTable("csv", "CSV", {in_csv}, {{"Scan All", "False"}})

    // Convert to a MEM table, which can be edited.
    view = ExportView(vw_csv + "|", "MEM", "IntermTripList", , )
    CloseView(vw_csv)
    // Create a unique iD
    tbl = CreateObject("Table", view)
    tbl.AddField("uniqueid_2")
    field_names = tbl.GetFieldNames()
    tbl = null
    v_tid = GetDataVector(view + "|", "uniqueid_2", )
    v_new_id = Vector(v_tid.length, "Long", {{"Sequence", start_id, 1}})
    SetDataVector(view + "|", "uniqueid_2", v_new_id,)

    // Determine which records translate to vehicles
    SetView(view)
    veh_set = CreateSet("vehicles")
    if mode_filter <> null
        then query = mode_filter + "and " + oriField + " <> " + desField
        else query = "Select * where " + oriField + " <> " + desField
    num_veh = SelectByQuery(veh_set, "several", query)
    if num_veh = 0 then Throw(
        "No vehicles found in the trip table. Query used:\n" +
        query
    )

    if field_names.position("finalDepartMinute") > 0 then do
        tbl = CreateObject("Table", view)
        tbl.RenameField({FieldName: "finalDepartMinute", NewName: "finalDeparture"})
        tbl = null
    end

    // Collect field data vectors
    pnum_col_name = if b_car_tracker then "driverId" else "pnum"
    get_fields = {
        oriField,
        desField,
        "finalDeparture",
        "uniqueid_2",
        "hhid",
        pnum_col_name,
        "origPurp",
        "destPurp"
    }
    if field_names.position("party") > 0 then get_fields = get_fields + {"party"}
    if field_names.position("activityMinutesAtDest") > 0 then get_fields = get_fields + {"activityMinutesAtDest"}
    opts = null
    opts.OptArray = "true"
    input = GetDataVectors(view + "|" + veh_set, get_fields, opts)
    num_records = input.(oriField).length
    data = null
    data.Trip = input.uniqueid_2
    data.Origin = input.(oriField)
    data.Destination = input.(desField)
    data.HH = input.hhid
    if b_car_tracker then
        data.DriverID = input.driverId
    else
        data.PersonNum = input.hhid * 100 + input.pnum
    data.OriginPurpose = input.origPurp
    data.DestinationPurpose = input.destPurp
    // TSM needs departure time in seconds after midnight.
    // Input departure times are in minutes after 3am
    data.DepartureTime = round(input.finalDeparture * 60, 1) + 10800 
    if field_names.position("party") > 0 then do
        // Determine occupancy. The 'party' field has values like:
        // "[ 1 ]"
        // "[ 4,5 ]"
        // Where the numbers are the person numbers that are in the party for the 
        // trip. Parse that to determine occupancy.
        v_len = StringLength(input.party)
        s = input.party
        s = Right(s, StringLength(s) - 2)
        s = Left(s, StringLength(s) - 2)
        dim occ[s.length]
        for i = 1 to s.length do
            temp = s[i]
            parts = ParseString(temp, ",")
            occ[i] = parts.length
        end
        data.Occupants = A2V(occ)
    end
    // Calculate previous trip number
    v_prevtrip = vector(input.uniqueid_2.length, "Long", )
    v_pnum = input.(pnum_col_name)
    for i = 1 to input.uniqueid_2.length do
        if i = 1 then do
            v_prevtrip[i] = null
            continue
        end
        person_num = v_pnum[i]
        prev_person_num = v_pnum[i - 1]
        if person_num <> prev_person_num then do
            v_prevtrip[i] = null
            continue
        end
        v_prevtrip[i] = data.Trip[i - 1]
        
    end
    data.PreviousTrip = v_prevtrip
    if field_names.position("activityMinutesAtDest") > 0 then do
        // The activity duration in the trip list can be negative, but doesn't
        // make sense for TSM.
        data.ActivityDuration = max(input.activityMinutesAtDest, 0)
    end
    // If the ValueOfTime column is null, it causes errors
    data.VOT = Vector(data.Trip.length, "float", {Constant: 0})

    // Create table
    pnum_field_name = if b_car_tracker then "DriverID" else "PersonNum"
    fields = {
        {"Trip",               "Real",     8, 0, "False"},
        {"Origin",             "Real",     8, 0, "False"},
        {"Destination",        "Real",     8, 0, "False"},
        {"DepartureTime",      "Real",     8, 2, "False"},
        {"PreviousTrip",       "Real",     8, 0, "False"},
        {"VOT",                "Real",     8, 2, "False"},
        {"HH",                 "Integer", 16, 0, "False"},
        {pnum_field_name,      "Integer", 16, 0, "False"},
        {"OriginPurpose",      "Integer",  8, 0, "False"},
        {"DestinationPurpose", "Integer",  8, 0, "False"}
    }
    if field_names.position("party") > 0 then fields = fields + {{"Occupants", "Real", 4, 0, "False"}}
    if field_names.position("activityMinutesAtDest") > 0 then fields = fields + {{"ActivityDuration", "Real", 8, 2, "False"}}
    vw_out = CreateTable("output", , "MEM", fields)

    opts = null
    opts.[Empty Records] = num_veh
    AddRecords(vw_out, , , opts)
    SetDataVectors(vw_out + "|", data, )

    // Per Matt Stratton's email on 1/31/2023, exclude IE and EI trips so that it's only II trips.
    SetView(vw_out)
    lowestMAZvalue = i2s(20001)
    highestTAZvalue = i2s(2186)
    if od_geography = "TAZ" then
        query = "Select * where Origin <= " + highestTAZvalue + " and Destination <= " + highestTAZvalue
    else
        query = "Select * where Origin >= " + lowestMAZvalue + " and Destination >= " + lowestMAZvalue
    num_IItrips = SelectByQuery("II", "several", query)
    opts = null
    opts.[CSV Header] = "true"
    opts.[Row Order] = {{"Trip", "Ascending"}}
    opts.[Indexed Fields] = {"Trip", "DepartureTime"}
    ExportView(vw_out + "|II", "CSV", out_csv, , opts)
    
    CloseView(view)
    CloseView(vw_out)
    return(True)
EndMacro


Macro "test convert DCOM trip list"
    on escape do return() end
    in_csv = ChooseFile({{"Trip List", "*.csv"}, {"Trip List", "*.bin"}}, "Choose a trip list file", )
    on escape default
    parts = SplitPath(in_csv)
    out_folder = ChooseDirectory("Choose the folder in which to save the converted trip list", {{"Initial Directory", parts[1] + parts[2]}})
    out_csv = out_folder + "\\" + parts[3] + parts[4]
    RunMacro("convert DCOM trip list", in_csv, out_csv )
    ShowMessage("Done")
endMacro

Macro "convert DCOM trip list" (in_csv, out_csv)

    if in_csv = null then Throw("'in_csv' not provided")
    if out_csv = null then Throw("'out_csv' not provided")

    /* Fields in DCOMVehicleTrips.csv
    EmpID       	    Employee ID
    estabType   	    1=Industry, 2=Wholesale, 3=Retail, 4=Tranport, 5=Service
    subEstabType	    There are 5 subestab under Service, and 4 subestab under Industry (see below from file dcomEmpCatbyEstablishment.csv)
    numOfTrip	        Number of trips in tour
    vehicleType	        Light, Medium, Heavy
    estabTaz	        Establishment TAZ
    currentTaz	        Trip Origin TAZ
    selectedTaz	        Trip Destination TAZ
    purpose	            0=STAY, 1=GOOD_PURPOSE, 2=SERVICE_PURPOSE, 3=OTHER_PURPOSE, 4=MEETING_PURPOSE, 5=BACK_TO_ESTAB
    beginTime	        minutes from midnight
    driveTime	        Drive time from currentTAZ to selectedTAZ
    driveDist	        Drive distance from currentTAZ to selectedTAZ
    period	            AM, MD, PM, NT 
    isGoodTour          Created to flag trips that should be discarded before writing out final trip list and trip tables. 1 = discard, 0 = keep. Should always be 0 in table we receive.
    isOtherMeetingTour  Created to flag trips that should be discarded before writing out final trip list and trip tables. 1 = discard, 0 = keep. Should always be 0 in table we receive.

    NAICS_CODE	EMP_CAT	                                                                    ESTAB	        SUB_ESTAB
    NAICS11	    Agriculture, Forestry, Fishing and Hunting	                                Industrial	    AgForFish
    NAICS21	    Mining	                                                                    Industrial	    OtherMfg
    NAICS22	    Utilities	                                                                Service	        Serv_Other
    NAICS23	    Construction	                                                            Industrial	    Construction
    NAICS31	    Manufacturing	                                                            Industrial	    OtherMfg
    NAICS32	    Manufacturing	                                                            Industrial	    OtherMfg
    NAICS33	    Manufacturing	                                                            Industrial	    HeavyMfg
    NAICS42	    Wholesale Trade	                                                            Wholesale	    Wholesale
    NAICS44	    Retail Trade	                                                            Retail	        Retail
    NAICS45	    Retail Trade	                                                            Retail	        Retail
    NAICS48	    Transportation and Warehousing	                                            Transportation	Transportation
    NAICS49	    Transportation and Warehousing	                                            Transportation	Transportation
    NAICS51	    Information	                                                                Service	        Serv_Other
    NAICS52	    Finance and Insurance	                                                    Service	        Serv_Other
    NAICS53	    Real Estate and Rental and Leasing	                                        Service	        Serv_Hotel
    NAICS54	    Professional, Scientific and Technical Services	                            Service	        Serv_Other
    NAICS55	    Management of Companies and Enterprises         	                        Service	        Serv_Other
    NAICS56	    Administrative and Support and Waste Management and Remediation Services	Service	        Serv_Other
    NAICS61	    Education Services	                                                        Service	        Serv_Education
    NAICS62	    Health Care and Social Assistance	                                        Service	        Serv_Health
    NAICS71	    Arts, Entertainment and Recreation	                                        Service	        Serv_Other
    NAICS72	    Accommodation and Food Services	                                            Service	        Serv_Hotel
    NAICS81	    Other Services (except Public Administration)	                            Service	        Serv_Other
    NAICS92	    Public Administration	                                                    Service	        Serv_Government
    NAICS99	    Other Unknown Employment	                                                Service	        Serv_Other

    The time period break points are:
    AM start minute = 360
    MD start minute = 540
    PM start minute = 900
    NT start minute = 1140
    NT end  minute = 1800
    */

    // Create a memory table of the input DCOM trip list
    vw_csv = OpenTable("csv", "CSV", {in_csv}, {{"Scan All", "False"}})
    SetView(vw_csv)
    // Remove any trips past 1440 (they never get simulated)
    SelectByQuery("trips", "several", "Select * where beginTime < 1440")
    vw = ExportView(vw_csv + "|trips", "MEM", "mem_tbl", , )
    CloseView(vw_csv)

    // The DCOM csv file does not have a unique ID field. Create a FFB copy of the table with a unique ID field created.
    // This facilitates joining operations after a simulation. And helps calculate Previous Trip.
    // Subsequent processes assume the DCOM list was sorted by EmpID and beginTime.
    num_records = GetRecordCount(vw, )
    unique_id = vector(num_records, "Long", {{"Sequence", 20000000, 1}})
    RunMacro("TCB Add View Fields", {vw, {{"unique_id", "integer"}}})
    SetDataVector(vw + "|", "unique_id", unique_id, )
    DCOMtrips_with_uid = GetTempFileName(".bin")
    ExportView(vw + "|", "FFB", DCOMtrips_with_uid, , )

    // Collect field data vectors
    get_fields = {
        "unique_id",    // Added above
        "EmpID",       	// Employee ID (tour ID)
        "currentTaz",   // Trip Origin TAZ
        "selectedTaz",  // Trip Destination TAZ
        "beginTime",	// minutes from midnight
        "driveTime",	// Drive time from currentTAZ to selectedTAZ
        "vehicleType"	// Light, Medium, Heavy
    }
    opts = null
    opts.OptArray = "true"
    input = GetDataVectors(vw + "|", get_fields, opts)
    data = null
    data.Trip = input.unique_id
    data.Origin = input.currentTaz
    data.Destination = input.selectedTaz
    // DCOM input departure times are in minutes after midnight
    // TSM needs seconds after midnight. Also, spread out trips uniformly over
    // the minute.
    data.DepartureTime = round(input.beginTime * 60, 1) + RandomNumber() * 60
    
    // Assume always 1 occupant for DCOM trips
    data.Occupants = vector(input.unique_id.length, "Short", {{"Constant", 1}})

    // Calculate previous trip number, activity duration, and vehicle class. Assumes DCOM list sorted by EmpID and beginTime
    num_trips = input.unique_id.length
    data.PreviousTrip = vector(num_trips, "Long", )
    data.ActivityDuration = vector(num_trips, "Long", )
    data.VehicleClass = vector(num_trips, "String", )

    last_indexDiffOrigDest = 0
    for i = 1 to num_trips do
        // Handle Previous Trip Number and Activity Duration when i = 1
        if i = 1 then do 
            data.ActivityDuration[i] = 0 // this will get reset when i = 2
            data.PreviousTrip[i] = null
            last_indexDiffOrigDest = i
        end
        // Handle if the current trip's origin and destination are the same. 
        // If the current trip's origin and destination are the same, then
        //    1. Recalculate activity duration of prior trip.
        //    2. set PrevTrip to 0
        else if data.Origin[i] = data.Destination[i] then do
            data.ActivityDuration[i] = -1 // a flag to identify which activity durations to ignore 
            data.PreviousTrip[i]     = -1 // a flag to identify which activity durations to ignore 
            goto skip_calculation
            end
        else do
            // Activity Duration of Prior Trip
            if input.EmpID[i] = input.EmpID[last_indexDiffOrigDest] then 
                data.ActivityDuration[last_indexDiffOrigDest] = max(0, input.beginTime[i] - input.beginTime[last_indexDiffOrigDest] - input.driveTime[last_indexDiffOrigDest])
            else
                data.ActivityDuration[last_indexDiffOrigDest] = null
            // Previous Trip Number
            if input.EmpID[i] = input.EmpID[last_indexDiffOrigDest] then 
                data.PreviousTrip[i] = data.Trip[last_indexDiffOrigDest]
            else 
                data.PreviousTrip[i] = null
            // Set count back to 0
            last_indexDiffOrigDest = i
        end
        // Vehicle Class
        if      input.vehicleType[i] = "TRKLT"  then data.VehicleClass[i] = "Light Truck"
        else if input.vehicleType[i] = "TRKMED" then data.VehicleClass[i] = "Medium Truck"
        else if input.vehicleType[i] = "TRKHVY" then data.VehicleClass[i] = "Heavy Truck"
        else    Throw("Invalid vehicle class in row " + i2s(i))
    skip_calculation:
    end

    // Create table
    fields = {
        {"Trip", "Real",  8, 0, "False"},
        {"Origin", "Real",  8, 0, "False"},
        {"Destination", "Real",  8, 0, "False"},
        {"DepartureTime", "Real",  8, 2, "False"},
        {"Occupants", "Real",  4, 0, "False"},
        {"PreviousTrip", "Real",  8, 0, "False"},
        {"ActivityDuration", "Real",  8, 2, "False"},
        {"VehicleClass", "String",  15, 2, "False"}
    }
    vw_out = CreateTable("output", , "MEM", fields)
    opts = null
    opts.[Empty Records] = num_trips
    AddRecords(vw_out, , , opts)
    SetDataVectors(vw_out + "|", data, )

    // Exclude trips that start and end at same TAZ.
    SetView(vw_out)
    trips_set = "valid trips"
    query = "Select * where ActivityDuration != -1 & PreviousTrip != -1"
    num_valid_trips = SelectByQuery(trips_set, "several", query)
    if num_valid_trips = 0 then Throw(
        "No valid trips found in the trip table. Query used:\n" +
        query
    )

    // Export and close views
    ExportView(vw_out + "|" + trips_set, "CSV", out_csv, , {[CSV Header]: "true"})
    CloseView(vw)
    CloseView(vw_out)
    return(True)
EndMacro

/*
The SMP must be open.
This macro creates the trajectory csv needed by CTRAMP. This macro is called automatically
at the end of every run. For details, see Simulation > Options... > Macros
*/

Macro "create trajectory csv"
    shared trips_with_uid

    // Collect all data needed for trajectories
    // Calculate travel time for each vehicle in the TransModeler output trip
    // table.
    rm = CreateObject("TSM.RunManager")
    output_file = rm.GetTripTable()
    if output_file = null then do
        opts = null
        opts.Title = "Choose Trip Data Table"
        opts.Help = "main|About_the_Path_Tools"
        opts.Table = "Trips"
        output_file = RunDbox("tsm choose output table", opts)
    end
    {dir, path, name, ext} = SplitPath(output_file)
    output_dir = dir + path
    output_vw = OpenTable("trips", "FFB", {output_file})
    fields = {
        "ID",
        "Class",
        "OriID",
        "DesID",
        "DepTime",
        "ArrTime",
        // "travel_time",
        "Status",
        "Path",
        "PathTable"
    }
    opts = null
    opts.OptArray = "true"
    input = GetDataVectors(output_vw + "|", fields, opts)
    input.travel_time = (input.ArrTime - input.DepTime) / 60

    // Create trajectory csv
    data = null
    data.vehicle_id = input.ID
    data.demand_type = if (input.Class = "PU" or input.Class = "PC1" or input.Class = "M")
        then "SOV"
        else if (input.Class = "PC1" or input.Class = "PC3")
            then "HOV"
            else "truck"
    data.origin_zone_id = input.OriID
    data.destination_zone_id = input.DesID
    data.departure_time_in_min = input.DepTime / 60
    data.arrival_time_in_min = input.ArrTime / 60
    data.travel_time_in_min = input.travel_time
    data.trip_complete_flag = if (input.Status = "Completed")
        then 1
        else 0
    data.path_node_sequence = String(input.Path)
    fields = {
        {"vehicle_id", "Integer", 8, 0, "True"},
        {"demand_type", "String", 8, 0, "True"},
        {"origin_zone_id", "Integer", 8, 0, "True"},
        {"destination_zone_id", "Integer", 8, 0, "True"},
        {"departure_time_in_min", "Float", 8, 2, "True"},
        {"arrival_time_in_min", "Float", 8, 2, "True"},
        {"travel_time_in_min", "Float", 8, 2, "True"},
        {"trip_complete_flag", "Integer", 8, 0, "True"},
        {"path_node_sequence", "String", 5000, 0, "True"}
    }
    vw_csv = CreateTable("table", , "MEM", fields)
    opts = null
    opts.[Empty Records] = data.vehicle_id.length
    AddRecords(vw_csv, , , opts)
    SetDataVectors(vw_csv + "|", data, )
    scen_name = rm.GetCurrentScenario()
    period = Left(scen_name, 2)

    // Convert the path IDs to sequence of nodes
    RunMacro("tsm import path table", null, {{"Trip Table", output_file}, {"Silent", "True"}})
    {path_table_indices, path_table_fnms, path_tables} = GetPathTables(output_file)
    
    for i = 1 to input.Path.length do
        path_id = input.Path[i]
        path_index = input.PathTable[i]
        info = GetPathInfo({path_id}, {{"Path Table", path_index}})
        link_ids = info[1][2]
        nodes = null
        for link_id in link_ids do
            if nodes = null then nodes = GetLinkNodes(link_id)
            else do
                temp = GetLinkNodes(link_id)
                nodes = nodes + {temp[2]}
            end
        end

        nodes = RunMacro("convert TM nodes to TC nodes", nodes)
        
        nodes = RunMacro("A2S", nodes)
        nodes = Substitute(nodes, "{", "", )
        nodes = Substitute(nodes, "}", "", )
        nodes = Substitute(nodes, ", ", ";", )
        
        data.path_node_sequence[i] = nodes
    end
    SetDataVector(vw_csv + "|", "path_node_sequence", data.path_node_sequence, )

    traj_file = output_dir + "trajectory_" + period + ".csv"
    ExportView(
        vw_csv + "|",
        "CSV",
        traj_file, ,
        {[CSV Header]: "true"}
    )

    CloseView(output_vw)
    CloseView(vw_csv)
    return(1)
endmacro

/*
Helper macro
*/

Macro "convert TM nodes to TC nodes" (tm_nodes)

    tc_node_layer = "2018 Nodes 190401"
    tc_link_layer = "2018 Links 190401"
    tm_set = CreateSet("tm node")
    for tm_node in tm_nodes do
        SetLayer("Nodes")
        SelectByIDs(tm_set, "several", {tm_node})
        SetLayer(tc_node_layer)
        opts = null
        opts.max = 1
        n = SelectNearestFeatures("tc node", "several", "Nodes|" + tm_set, 100/5280, opts)
        if n = 0 then continue
        tc_id = GetSetIDs("tc node")
        tc_nodes = tc_nodes + tc_id
    end

    if tc_nodes = null then return({-99})
    if tc_nodes.length < 2 then return({-99})
    SetLayer(tc_link_layer)
    temp_net = GetTempFileName(".net")
    net = CreateNetwork(, temp_net, "temp", , , )
    short_path = ShortestPath(net, tc_nodes, 1, )
    path_length = short_path[1]
    link_ids = short_path[2]
    link_dirs = short_path[3]
    SetLayer(tc_link_layer)
    for i = 1 to link_ids.length do
        link_id = link_ids[i]
        link_dir = link_dirs[i]

        endpoints = GetEndPoints(link_id)
        if link_dir = 1 then endpoints = {endpoints[2], endpoints[1]}
        if final_nodes = null then final_nodes = endpoints
        else final_nodes = final_nodes + {endpoints[2]}
    end

    return(final_nodes)
endmacro

/*
Divides the original departure bins into smaller, sub bins. Importantly, the
probability for sub bins is linearly interpolated between adjacent bins.

i.e.
      __
 __  |  |
|  | |  |
|  | |  |

becomes something like:

       _
   _ _| |
 _| | | |
| | | | |

Inputs
  * view
    * string
    * view name to edit
  * sub_bins
    * integer
    * number of sub bins to create

*/

Macro "interpolate departure times" (view, veh_set)

    // The number of bins to create from each original bin. The code below only
    // works if set to 3. It is not general enough to work with other numbers.
    sub_bins = 3

    // Aggregate the view to determine the departure rate per 30 minute bin
    temp_bin = GetTempFileName(".bin")
    vw_ag = AggregateTable(
        "agg", view + "|" + veh_set, "FFB", temp_bin, "prelimDepartInterval",
        {{"prelimDepartInterval", "COUNT"}},
    )
    v_prelimDepartInterval = GetDataVector(vw_ag + "|", "prelimDepartInterval", )
    v_count = GetDataVector(vw_ag + "|", "N prelimDepartInterval", )
    CloseView(vw_ag)

    // If we just split the bins up evenly, these would be the number of trips
    // per sub bin.
    v_sub_bin_count = v_count / sub_bins

    // Calculate left and right mid points (between bins)
    dim a_left_mp[v_sub_bin_count.length]
    dim a_right_mp[v_sub_bin_count.length]
    for i = 1 to v_sub_bin_count.length do
        mid_bin = v_sub_bin_count[i]
        previous_mid_bin = if i = 1
            then 0
            else v_sub_bin_count[i - 1]
        next_mid_bin = if i = v_sub_bin_count.length
            then 0
            else v_sub_bin_count[i + 1]
        
        a_left_mp[i] = (mid_bin - previous_mid_bin) / 2 + previous_mid_bin
        a_right_mp[i] = (next_mid_bin - mid_bin) / 2 + mid_bin
    end
    
    // Assign each trip to a departure sub interval
    a_fields = {{"dep_sub_interval", "Integer", 10}}
    RunMacro("Add Fields", view, a_fields, 0)
    for i = 1 to v_prelimDepartInterval.length do
        prelimDepartInterval = v_prelimDepartInterval[i]
        mid_bin = v_sub_bin_count[i]
        left_mp = a_left_mp[i]
        right_mp = a_right_mp[i]

        SetView(view)
        n = SelectByQuery(
            "set", "several", 
            "Select * where prelimDepartInterval = " + String(prelimDepartInterval),
            {"Source And": veh_set}
        )
        if n = 0 then continue

        new_bin = null
        for j = 1 to sub_bins do
            new_bin = new_bin + {prelimDepartInterval * sub_bins - sub_bins + j}
        end
        weight = {left_mp, mid_bin, right_mp}
   
        opts = null
        opts.population = new_bin
        opts.weight = weight
        result = RandSamples(n, "Discrete", opts)
        SetDataVector(view + "|set", "dep_sub_interval", result, )
    end
endmacro

/*
This script is used to quickly aggregate simulation volumes into periods to
make it easy to create scatter plots or do other comparisons.
*/

Macro "aggregatevols" 
    shared d_edit_options
	RunManager = CreateObject("TSM.RunManager")
	scenario = RunManager.GetCurrentScenario()
    folder = RunManager.GetOutputFolder()

	// Run Simulation
    RunManager.SuppressAllWarnings()
    RunManager.SetSimulationRunMode("Simulation") 
    RunManager.SetSubareaAnalysis("False", ) 
    RunManager.SetAggregateVolumes("True", 15, ) 
    RunManager.MinimizeSimulationWindows()
    RunManager.RunSimulation()

    // Add expression fields
    expressionAM_AB = "AB_Vol_0600 + AB_Vol_0615 + AB_Vol_0630 + AB_Vol_0645 + AB_Vol_0700 + AB_Vol_0715 + AB_Vol_0730 + AB_Vol_0745 + AB_Vol_0800 + AB_Vol_0815 + AB_Vol_0830 + AB_Vol_0845"
    expressionAM_BA = "BA_Vol_0600 + BA_Vol_0615 + BA_Vol_0630 + BA_Vol_0645 + BA_Vol_0700 + BA_Vol_0715 + BA_Vol_0730 + BA_Vol_0745 + BA_Vol_0800 + BA_Vol_0815 + BA_Vol_0830 + BA_Vol_0845"
    expressionPM_AB = "AB_Vol_1400 + AB_Vol_1415 + AB_Vol_1430 + AB_Vol_1445 + AB_Vol_1500 + AB_Vol_1515 + AB_Vol_1530 + AB_Vol_1545 + AB_Vol_1600 + AB_Vol_1615 + AB_Vol_1630 + AB_Vol_1645 + AB_Vol_1700 + AB_Vol_1715 + AB_Vol_1730 + AB_Vol_1745"
    expressionPM_BA = "BA_Vol_1400 + BA_Vol_1415 + BA_Vol_1430 + BA_Vol_1445 + BA_Vol_1500 + BA_Vol_1515 + BA_Vol_1530 + BA_Vol_1545 + BA_Vol_1600 + BA_Vol_1615 + BA_Vol_1630 + BA_Vol_1645 + BA_Vol_1700 + BA_Vol_1715 + BA_Vol_1730 + BA_Vol_1745"

    if scenario contains "AM" then do
        expression_AB = expressionAM_AB
        expression_BA = expressionAM_BA
        expressionname_AB = "AB_Vol_0600-0900"
        expressionname_BA = "BA_Vol_0600-0900"
    end
    else do
        expression_AB = expressionPM_AB
        expression_BA = expressionPM_BA
        expressionname_AB = "AB_Vol_1400-1800"
        expressionname_BA = "BA_Vol_1400-1800"
    end

    viewSegmentVolumeSpeeds = "Segment Volumes & Speeds"
    createexpr_AB = CreateExpression(viewSegmentVolumeSpeeds, expressionname_AB, expression_AB, )
    createexpr_BA = CreateExpression(viewSegmentVolumeSpeeds, expressionname_BA, expression_BA, )
    
    // join Segments and Aggregate Volumes on segment ID
	viewSegments = "Segments"
    viewJoin = JoinViews("Segments + Segment Flows", viewSegments+".ID", viewSegmentVolumeSpeeds+".ID",  )
    editor = CreateEditor("Segments + Segment Flows", viewJoin + "|", , d_edit_options)

endMacro

/*
This macro runs a single simulation to produce the output needed to create dynamic skims
and then creates the dynamic skim matrices.
*/

Macro "Create Dynamic Skims" (RunMgr)
    if RunMgr = null then RunMgr = CreateObject("TSM.RunManager") // called independently to test

    // first run a simulation to produce the dynamic skim data
    RunMgr.SetSimulationRunMode("Simulation")
    RunMgr.SetDynamicSkims("True")
    RunMgr.MinimizeSimulationWindows()
    RunMgr.RunSimulation()

    folder = RunMgr.GetScenarioFolder()
    skimRuns = RunMgr.GetDynamicSkimRuns()
    skim_file = folder + "dynamic_skim.mtx"
    skimOpts = {
        {"Filename"   , skim_file},
        {"Run"        , skimRuns[skimRuns.length]},
        {"Matrix Type", "Dynamic"           },
        {"Data"       , "Travel Time"       },
        {"Interval"   , 15                  },
        {"Square"     , True                },
        {"Estimate"   , True}
    }
    mtxHandles = RunMgr.CreateDynamicSkimMatrix(skimOpts)

    mtx = CreateObject("Matrix", skim_file)
    core_names = mtx.GetCoreNames()
    mtx.AddCores("diff")
    mtx.AddCores("FFTime")
    // FFTime is the minimum of all the core times
    for i = 1 to core_names.length do
        core = core_names[i]
        if core = "diff" or core = "FFTime" then continue

        if i = 1 then mtx.FFTime := mtx.(core)
        else mtx.FFTime := min(mtx.FFTime, mtx.(core))
    end    
    for core in core_names do
        if core = "diff" or core = "FFTime" then continue

        temp = Substitute(core, "Time_", "", )
        hour = S2I(Left(temp, 2))
        minute = S2I(Right(temp, 2))

        // Null out cores between 10pm and 5am
        if hour >= 22 or hour <= 5 then do
            mtx.(core) := null
            continue
        end

        // For all remaining cores, null out any times <5% different from FFTime
        mtx.diff := abs(mtx.(core) - mtx.FFTime) / mtx.FFTime
        mtx.(core) := if mtx.diff < 0.05 then null else mtx.(core)
    end
    mtx.DropCores("diff")

    // Copy the matrix which reduces the size
    core_names = mtx.GetCoreNames()
    CopyMatrix(mtx.(core_names[1]), {
        "File Name": folder + "dynamic_skim_small.mtx",
        Label: "Dynamic Skim Small",
        Compression: 1
    })
    mtx = null
    mtxHandles = null
    DeleteFile(skim_file)
EndMacro





















/*doc
Adds a field.  Replacement for hidden "TCB Add View Fields", which has
some odd behavior.  Takes the same field info array.  See ModifyTable()
for the 12 potential elements.
  * view
    * String
    * view name
  * a_fields
    * Array of arrays
    * Each sub-array contains the 12-elements that describe a field.
      e.g. {{"Density", "Real", 10, 3, , , , "Used to calculate initial AT"}}
      (See ModifyTable() TC help page for full array info)
  * initial_values
    * Number, string, or array of numbers/strings (optional)
    * If not provided, any fields to add that already exist in the table will not be
      modified in any way. If provided, the added field will be set to this value.
      This can be used to ensure that a field is set to null, zero, etc. even if it
      already exists.
    * If a single value, it will be used for all fields.
    * If an array shorter than number of fields, the last value will be used
      for all remaining fields.
*/

Macro "Add Fields" (view, a_fields, initial_values)

  // Argument check
  if view = null then Throw("'view' not provided")
  if a_fields = null then Throw("'a_fields' not provided")
  for field in a_fields do
    if field = null then Throw("An element in the 'a_fields' array is missing")
  end
  if initial_values <> null then do
    if TypeOf(initial_values) <> "array" then initial_values = {initial_values}
    if TypeOf(initial_values) <> "array"
      then Throw("'initial_values' must be an array")
  end

  // Get current structure and preserve current fields by adding
  // current name to 12th array position
  a_str = GetTableStructure(view)
  for s = 1 to a_str.length do
    a_str[s] = a_str[s] + {a_str[s][1]}
  end
  for f = 1 to a_fields.length do
    a_field = a_fields[f]

    // Test if field already exists (will do nothing if so)
    field_name = a_field[1]
    exists = "False"
    for s = 1 to a_str.length do
      if a_str[s][1] = field_name then do
        exists = "True"
        break
      end
    end

    // If field does not exist, create it
    if !exists then do
      dim a_temp[12]
      for i = 1 to a_field.length do
        a_temp[i] = a_field[i]
      end
      a_str = a_str + {a_temp}
    end
  end

  ModifyTable(view, a_str)

  // Set initial field values if provided
  if initial_values <> null then do
    nrow = GetRecordCount(view, )
    for f = 1 to a_fields.length do
      field = a_fields[f][1]
      type = a_fields[f][2]
      if f > initial_values.length
        then init_value = initial_values[initial_values.length]
        else init_value = initial_values[f]

      if type = "Character" then type = "String"

      opts = null
      opts.Constant = init_value
      v = Vector(nrow, type, opts)
      SetDataVector(view + "|", field, v, )
    end
  end
EndMacro

Macro "Setup DTA"
    SetSimulationProfileItems("Link", {{"DestinationZone", "Zone"}})
endMacro


/*
Converts arrays to strings. An array of {1,2,3} will be
converted to "{1,2,3}".
*/

Macro "A2S" (array)

  if TypeOf(array) <> "array" then Throw("'array' must be an array")

  for a = 1 to array.length do
    temp = array[a]

    if TypeOf(temp) = "matrix" then temp = RunMacro("M2A", temp)
    if TypeOf(temp) = "array" then temp = RunMacro("A2S", temp)
    else if TypeOf(temp) <> "string" then temp = String(temp)

    if a = 1 then string = string + "{" + temp
    else string = string + ", " + temp
  end
  string = string + "}"

  return(string)
EndMacro