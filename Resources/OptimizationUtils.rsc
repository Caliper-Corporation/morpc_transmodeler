Macro "UpdateOptMinTimes"
    // per lane min green times
    MINGREEN_L  = 4
    MINGREEN_TR = 5

    runManager = RunMacro("CreateRunManager")
    nodeLayer = runManager.GetLayerName("Node")
    signalizedNodeIDs = GetSetIDs(nodeLayer + "|Signalized")
    //signalizedNodeIDs = {23412}
    opts = null
    opts.Nodes = signalizedNodeIDs
    signalPlansFile = runManager.GetSignals()
    opts.SignalPlans = signalPlansFile
    simulationPeriod = runManager.GetSimulationPeriod()
    simulationStartTime = SecondsToString(simulationPeriod[1])
    opts.StartTime = simulationPeriod[1]
    SetCursor("Hourglass")
    on error do
        ShowMessage(GetLastError())
        return()
        end
    icInfo = GetIcNodesInfo(opts)
    ResetCursor()
    n = icInfo.length
    progressWindow = CreateObject("G30 Progress Window", 1, "Status")
    pbar = CreateObject("G30 Progress Bar", "Updating green time lower bounds...", true, n)
    for i = 1 to n do
        if pbar.Step() then
            goto cancelthismacro
        nodeIDs = icInfo[i].ids
        for j = 1 to nodeIDs.length do
            node = nodeIDs[j]
            lanes = GetNodeApproachLanes(node)
            // GetControllerPhases() only returns information about protected movements
            controllerPhases = GetControllerPhases(node, simulationStartTime, {SignalPlans: signalPlansFile})
            for k = 1 to controllerPhases.length do
                if controllerPhases[k].PedExclusive then continue
                turns = SortArray(controllerPhases[k].Turns, {Unique: "True"})
                bLeftTurnPhase = if turns.length = 1 & turns[1] = "L" then true else false
                upstreamLinks = controllerPhases[k].[Upstream Links]
                numLanes = 0 // number of lanes for the L (if bLeftTurnPhase) or for T, R, or shared (if not bLeftTurnPhase)
                for l = 1 to lanes.length do
                    reverseLane = GetReversibleLane(lanes[l])
                    movements = GetLaneMovements(lanes[l])
                    segment = GetLaneSegment(lanes[l])
                    link = GetSegmentLink(segment)
                    if ArrayPosition(upstreamLinks, {link},) > 0 then do
                        priority = GetLinkInfo(link).Priority
                        if bLeftTurnPhase then do
                            if reverseLane != null & movements contains "L" or movements = "L" then
                                numLanes = numLanes + 1
                            end
                        else if reverseLane != null then // include all lanes, whether L, T, or R because all are protected in this phase
                            numLanes = numLanes + 1
                        end
                    end
                if numLanes = 0 then continue
                priorityMultiplier = if priority <= 5 then 2 else   // major or minor arterial
                                     if priority <= 6 then 1.5 else // collector
                                     1                              // local
                minGreen = if bLeftTurnPhase then MINGREEN_L * numLanes else MINGREEN_TR * numLanes
                minGreen = Min(40, r2i(minGreen * priorityMultiplier))
                controllerPhases[k].OptMinGreen = minGreen
                end
            SetControllerPhases(node, simulationStartTime, controllerPhases, {SignalPlans: signalPlansFile})
            end
        end
cancelthismacro:
    pbar.Destroy()
endMacro