Attribute VB_Name = "VisioExport"
'==============================================================================
' RIDESIM EXPORTER FOR VISIO
'------------------------------------------------------------------------------
' Reads the ACTIVE PAGE and emits the three JSON blocks the Ride Sequence
' Planner expects (nodes, connections, attractions). Walking lines are read as
' polylines so their real bent shape drives length + animation.
'
' HOW TO USE
'   1. Alt+F11 in Visio -> File > Import File... > pick VisioExport.bas
'      (or paste this whole module into a new Module).
'   2. Build your drawing with these core masters named exactly:
'         Node       - a path junction (isAttraction = false)
'         Entrance   - an attraction entrance node (isAttraction = true)
'         Exit       - an attraction exit node     (isAttraction = true)
'         Attraction - the attraction label/marker (carries name + ride length)
'      Plus optional attraction-category masters - same as Attraction but tagged
'      with that category (no Entrance/Exit needed; link to one or more Nodes):
'         Restaurant / Shop / Pin / Restroom / Other
'      "Other" is a generic timed stop (default 5-min dwell from DEFAULT_RIDE;
'      override per shape with Prop.RideDuration).
'         Station    - a transit boarding stop (a graph node). Carries the
'                      boarding wait via Prop.ThPWID (live) and/or Prop.AvgWait.
'                      Id comes from Prop.ID, else its text (e.g. "Main St").
'   3. Wire it up with glued lines (center glue point to center glue point):
'        - WALKWAY graph: lines between Node / Entrance / Exit / Station shapes.
'          Bend them however you like - bends are captured as polylines.
'        - ASSOCIATION: from each Attraction shape, draw one line to its
'          Entrance node and one line to its Exit node. These tell the exporter
'          which entrance/exit belong to the attraction and are NOT walkway
'          edges. An Attraction must connect ONLY to Entrance/Exit shapes -
'          never directly into the walkway graph.
'        - TRANSPORT lines (railroad / ferries): put each line's tracks on a
'          layer named "transit_<LineName>" (e.g. transit_Railroad). Draw one
'          DIRECTED line per segment, carrying Prop.Duration = minutes for that
'          hop; its bends are captured as the segment path.
'          Direction: Visio won't glue a 2-D Station, so instead give the segment
'          two cells whose FORMULAS reference the stations - FromLink (departing)
'          and ToLink (arriving). Any of these cell homes works (first found):
'            Connections.FromLink.X / .Y, User.FromLink, or Prop.FromLink
'          (same for ToLink). The exporter reads the shape each formula points at
'          - just reference the station anywhere in the formula (e.g. =Station!PinX).
'          Prop.Reverse = TRUE flips the captured path order, so you can COPY a
'          one-way segment, swap its From/To links, and set Reverse instead of
'          redrawing the geometry backwards.
'          The engine chains segments in their direction:
'          - two-way: a segment each way (A->B and B->A; paths may differ).
'          - one-way loop: a cycle of segments (A->B, B->C, C->A).
'          Boarding wait comes from the BOARD station's ThPWID/AvgWait.
'   4. Run:  Tools > Macros > VisioExport.ExportRideSim   (or F5 in the editor).
'   5. The macro writes the data into the park's data file (park.js), replacing
'      the text between the // @RIDESIM:*:START / :END markers (the SAMPLE arrays).
'      The target is <docfolder>\parks\<active-page-slug>\park.js - i.e. one
'      folder per park, named after the Visio PAGE (e.g. page "Epcot" -> the
'      file parks\epcot\park.js). Set HTML_PATH_OVERRIDE below to force a path.
'      The park folder + its park.js (with the markers and a SAMPLE.meta block)
'      must already exist; just refresh the browser afterwards.
'      If park.js / its markers aren't found, it falls back to writing
'      "ridesim_export.txt" for manual paste. Output also goes to Ctrl+G.
'
' AUTO-NAMING (no shape data needed - just set shape TEXT):
'   Attraction : id = sanitized attraction TEXT (e.g. "Pirates" -> pirates).
'   Entrance   : id = <attractionId>_in   (taken from its association line).
'   Exit       : id = <attractionId>_out.
'   Node       : set its TEXT to the LAND it sits in (e.g. "Adventureland").
'                id = <land><n>, numbered left->right then bottom->top
'                (e.g. adventureland1, adventureland2 ...). Blank text -> node1..
'                A Node whose TEXT is "Start" gets id "start" and is used as the
'                planner's starting location (where the day begins).
'   Ids are forced unique; a clash gets _2, _3 appended.
'
' OPTIONAL SHAPE DATA (Shape Data / Custom Properties) - all override the above:
'   On any shape:        Prop.ID          -> use this exact id instead of auto
'                        Prop.Name        -> JSON "name" (else omitted for nodes)
'                        Prop.LatLon      -> "lat,lon" (e.g. "28.4177,-81.5812").
'                        Put it on 3+ spread-out shapes to calibrate GPS: the
'                        planner fits a lat/lon -> map-pixel transform and can
'                        show your live location + start from the nearest node.
'   On Attraction:       Prop.RideDuration-> ride minutes (else DEFAULT_RIDE).
'                        Also accepts Prop.Duration / RideTime / Ride / Minutes,
'                        numeric or text ("12 min").
'                        Prop.ThPWID      -> ThemeParks.wiki entity GUID, written
'                        as "thpwId"; the planner matches live standby waits +
'                        Lightning Lane to this ride by it. Aliases:
'                        ThpwId/ThemeParksID/TPWikiID.
'                        Prop.WaitID      -> (legacy) Queue-Times.com ride id,
'                        written as "waitId".
'                        Prop.AvgWait     -> typical wait (minutes), written as
'                        "avgWait"; used for sequence timing when live waits
'                        aren't on. Aliases: AverageWait/AvgWaitTime/WaitMinutes.
'   On any attraction:   Prop.Hovertext   -> extra detail shown on map hover
'                        (multi-line OK), written as "hoverText".
'                        Prop.Closed      -> true = drawn gray (not open today).
'                        Prop.Audio       -> URL/file looped while the avatar is
'                        at this stop during the animation, written as "audio".
'                        Prop.QSun        -> fraction (0..1, or a % like 40) of
'                        the QUEUE in the sun/heat, assumed at the front of the
'                        line; written as "qSun". Prop.QInside (legacy bool) still
'                        works and means a fully shaded queue (qSun = 0).
'                        Prop.RInside (aka "Indoor") -> true if the time at this
'                        stop is indoor/AC, false if outdoor. Feeds the sun-vs-AC
'                        bar and the marker heat ring (yellow = outdoor/false,
'                        blue = indoor/true) for restaurants & shops. Emitted as
'                        true or false when set; if unset, shops default indoor
'                        and everything else outdoor.
'
' SCALE BAR (sets real-world walking speed):
'   Add two shapes named (or mastered) ScaleStart and ScaleEnd and a line
'   between them; put the real distance in FEET as the line's text (e.g. "350").
'   The exporter writes feet-per-pixel = feet / pixel-distance into
'   SAMPLE.feetPerPixel, which fills the planner's ft/px box. (Distance text may
'   instead sit on ScaleStart/ScaleEnd if you prefer.)
'
' BACKGROUND MAP
'   Preferred: draw the map on its own page, set it as this page's BACKGROUND
'   (Design > Backgrounds, or right-click the page tab), and size both pages the
'   same. On export the background page is written to parks\<page-slug>\
'   background.svg automatically and SAMPLE.mapExtent becomes the whole page, so
'   the backdrop and the nodes always share one coordinate space (can't drift).
'   Run the macro from EITHER page: from the data page it exports its background;
'   from the map (background) page it exports that map and hops to the data page.
'   Legacy: put a map image on a layer named "Bck" and draw nodes over it; the
'   exporter writes that shape's bounding box as SAMPLE.mapExtent and you supply
'   background.png/.svg yourself. A "Bck" shape, if present, wins over the page.
'
' RIDE TRACKS (optional, for "interesting" rides):
'   Draw the ride path as a line/freeform shape on a layer named "Track", and
'   glue one end into the ride's Attraction shape (a connection point). The
'   exporter writes its vertices as "track" on that ride; the planner animates
'   the marker along it over the ride duration and draws it faintly on the map.
'   Vertices keep their natural order (first point = animation start).
'
' COORDINATES
'   Visio is inches, origin bottom-left, Y up. We emit pixels, origin top-left,
'   Y down (screen space). Node coords are relative to the Bck map extent, so
'   PPI only sets overall resolution (any value works; Visio default = 96).
'==============================================================================
Option Explicit

Private Const PPI As Double = 96#          ' pixels per inch (match your PNG export DPI)
Private Const DEFAULT_RIDE As Double = 5#  ' fallback ride duration (minutes)
' Path.Points flatness (inches): max gap between the approximating segments and the
' true curve. Smaller = smoother curves / more points (0.005" ~= 0.5px). At 0.05 a
' small ride circle flattened to ~8 sides (an octagon); this keeps curves smooth
' even zoomed in. Only curved spans densify — straight runs stay two points.
Private Const TRACK_FLATNESS As Double = 0.005
Private Const OUT_FILE As String = "ridesim_export.txt"
' Full path to the park data file (park.js) to patch. Leave "" to derive it
' from the saved Visio doc + active page name: <docfolder>\parks\<page-slug>\park.js
' (one folder per park, page name slugified -> folder name).
Private Const HTML_PATH_OVERRIDE As String = ""
' Layer holding the background map image; its extent defines where the
' background.png is stretched in node coordinates (nodes sit inside it).
Private Const BG_LAYER As String = "Bck"
' Layer holding ride track shapes. A shape on this layer is a ride animation
' path; glue one end into the ride shape to bind it to that ride.
Private Const TRACK_LAYER As String = "Track"
' Layer holding queue-line shapes: a per-ride path the avatar follows while
' waiting (entrance -> load). Glue one end into the ride, same as a Track.
Private Const QUEUE_LAYER As String = "Queue"
' Transport lines live on layers named "transit_<LineName>" (e.g. transit_Railroad).
Private Const TRANSIT_PREFIX As String = "transit_"

Private mPageH As Double                   ' active page height (inches), for Y-flip
Private mNodeMap As Collection             ' "k"&Shape.ID -> Array(id, cxPx, cyPx, role)
Private mAttrMap As Collection             ' "k"&Shape.ID -> attractionId
Private mRole As Collection                ' "k"&Shape.ID -> role (node-like shapes)
Private mScaleIds As Collection            ' "k"&Shape.ID of ScaleStart/ScaleEnd shapes
Private mTrackIds As Collection            ' "k"&Shape.ID of Track-layer shapes
Private mQueueIds As Collection            ' "k"&Shape.ID of Queue-layer shapes
Private mDataPage As Visio.Page            ' resolved data (foreground) page for this export
Private mWarnings As Collection

'------------------------------------------------------------------------------
Public Sub ExportRideSim()
    Dim pg As Visio.Page
    Set pg = Visio.ActivePage
    If pg Is Nothing Then MsgBox "No active page.", vbExclamation: Exit Sub

    ' Run from either page. If the active page is a background page it's the map —
    ' hop to the data page that uses it. Otherwise the active page is the data page
    ' and its assigned background (if any) is the map.
    Dim mapPage As Visio.Page
    If pg.Background <> 0 Then
        Set mapPage = pg
        Set pg = ForegroundUsing(pg)
        If pg Is Nothing Then MsgBox "This background (map) page isn't assigned to any data page." & vbCrLf & _
            "Open the data page, or set this page as its background, then re-run.", vbExclamation, "RideSim Export": Exit Sub
    Else
        Set mapPage = BackgroundPageOf(pg)
    End If
    Set mDataPage = pg

    mPageH = pg.PageSheet.CellsU("PageHeight").ResultIU
    Set mNodeMap = New Collection
    Set mAttrMap = New Collection
    Set mRole = New Collection
    Set mScaleIds = New Collection
    Set mTrackIds = New Collection
    Set mQueueIds = New Collection
    Set mWarnings = New Collection

    Dim shp As Visio.Shape, v As Variant
    Dim nodesJson As String, attrJson As String, connJson As String
    Dim nodeCount As Long, attrCount As Long, connCount As Long

    ' --- pass A: classify shapes; assign attraction ids ---------------------
    Dim nodeShapes As Collection, entShapes As Collection
    Dim exitShapes As Collection, attractions As Collection
    Set nodeShapes = New Collection
    Set entShapes = New Collection
    Set exitShapes = New Collection
    Set attractions = New Collection
    Dim usedAttr As Collection: Set usedAttr = New Collection
    Dim trackShapes As Collection: Set trackShapes = New Collection
    Dim queueShapes As Collection: Set queueShapes = New Collection
    Dim stationShapes As Collection: Set stationShapes = New Collection   ' transit Stop nodes
    Dim transitSegs As Collection: Set transitSegs = New Collection       ' directed track lines on transit_* layers

    For Each shp In pg.Shapes
        Dim role As String: role = MasterRole(shp)
        If role <> "" Then mRole.Add role, "k" & shp.id
        If ScaleRole(shp) <> "" Then mScaleIds.Add True, "k" & shp.id
        If OnLayer(shp, TRACK_LAYER) Then mTrackIds.Add True, "k" & shp.id: trackShapes.Add shp
        If OnLayer(shp, QUEUE_LAYER) Then mQueueIds.Add True, "k" & shp.id: queueShapes.Add shp
        If IsTransitLayer(shp) And role = "" Then transitSegs.Add shp   ' a transit track segment (any dimensionality)
        Select Case role
            Case "Node":     nodeShapes.Add shp
            Case "Station":  nodeShapes.Add shp: stationShapes.Add shp   ' a Stop is a graph node + a transit stop
            Case "Entrance": entShapes.Add shp
            Case "Exit":     exitShapes.Add shp
            Case "Attraction", "Restaurant", "Shop", "Pin", "Restroom", "Other"  ' all are attractions; category comes from the master
                mAttrMap.Add AttrIdFor(shp, usedAttr), "k" & shp.id
                attractions.Add shp
        End Select
    Next

    ' --- pass B: scan lines. A line touching an Attraction is an association
    '   (which Entrance/Exit belongs to it) and is NOT a walkway edge. Lines
    '   between graph nodes (Node/Entrance/Exit) are walkway edges. ---------
    Dim edgeLines As Collection: Set edgeLines = New Collection
    Dim entOf As Collection, exitOf As Collection
    Set entOf = New Collection   ' "k"&EntranceShapeID -> attrId
    Set exitOf = New Collection  ' "k"&ExitShapeID     -> attrId
    Dim directOf As Collection: Set directOf = New Collection ' "k"&AttractionShapeID -> node ShapeID (non-ride single-node link)
    For Each shp In pg.Shapes
        If MasterRole(shp) = "" And shp.OneD And Not KeyExists(mTrackIds, "k" & shp.id) And Not KeyExists(mQueueIds, "k" & shp.id) Then
          If IsTransitLayer(shp) Then
            ' a transit segment - collected in pass A; excluded from walkway edges here
          Else
            Dim bShp As Visio.Shape, eShp As Visio.Shape
            Set bShp = Nothing: Set eShp = Nothing
            GetEnds shp, bShp, eShp
            If IsScaleLine(bShp, eShp) Then
                ' the scale bar's line - measured separately, not a walkway edge
            ElseIf bShp Is Nothing Or eShp Is Nothing Then
                Warn "Line not glued at both ends, skipped: " & shp.NameU
            Else
                Dim kb As Visio.Shape, ke As Visio.Shape
                Set kb = ClimbKnown(bShp): Set ke = ClimbKnown(eShp)
                If kb Is Nothing Or ke Is Nothing Then
                    Warn "Line endpoint is not a known shape, skipped: " & shp.NameU
                ElseIf RoleOfShape(kb) = "Attraction" Or RoleOfShape(ke) = "Attraction" Then
                    Dim aShp As Visio.Shape, oShp As Visio.Shape
                    If RoleOfShape(kb) = "Attraction" Then
                        Set aShp = kb: Set oShp = ke
                    Else
                        Set aShp = ke: Set oShp = kb
                    End If
                    Dim aId As String: aId = mAttrMap("k" & aShp.id)
                    Select Case RoleOfShape(oShp)
                        Case "Entrance": PutOnceKey entOf, "k" & oShp.id, aId, "Entrance", shp
                        Case "Exit":     PutOnceKey exitOf, "k" & oShp.id, aId, "Exit", shp
                        Case "Attraction": Warn "Line links two Attractions, ignored: " & shp.NameU
                        Case "Node"
                            ' Rides need an Entrance + Exit; everything else
                            ' (restaurant/shop/pin) hooks to one OR MORE nodes -
                            ' the planner routes to whichever is nearest.
                            If CategoryOf(aShp) = "ride" Then
                                Warn "Ride '" & aId & "' linked to a plain Node - rides use Entrance/Exit. Ignored."
                            Else
                                Dim lst As Collection
                                If KeyExists(directOf, "k" & aShp.id) Then
                                    Set lst = directOf("k" & aShp.id)
                                Else
                                    Set lst = New Collection: directOf.Add lst, "k" & aShp.id
                                End If
                                lst.Add CStr(oShp.id)
                            End If
                        Case Else: Warn "Attraction '" & aId & "' linked to an unexpected shape. Ignored."
                    End Select
                ElseIf kb.id <> ke.id Then
                    edgeLines.Add shp           ' walkway edge; ids resolved in pass D
                Else
                    Warn "Line skipped (both ends on same shape): " & shp.NameU
                End If
            End If
          End If
        End If
    Next

    ' --- pass C: assign node ids -------------------------------------------
    '   Entrance/Exit -> <attractionId>_in / _out. Plain Nodes -> <land><n>,
    '   numbered in spatial order (left->right, then bottom->top).
    Dim usedNode As Collection: Set usedNode = New Collection
    Dim assocEnt As Collection, assocExit As Collection
    Set assocEnt = New Collection   ' attrId -> entrance nodeId
    Set assocExit = New Collection  ' attrId -> exit nodeId

    Dim ordered As Collection: Set ordered = SortNodesSpatially(nodeShapes)
    Dim landCount As Collection: Set landCount = New Collection
    For Each v In ordered
        Set shp = v
        Dim land As String: land = Sanitize(ShapeText(shp))
        Dim nid As String
        Dim nrole As String: nrole = "Node"
        If RoleOfShape(shp) = "Station" Then     ' a transit Stop: id from Prop.ID, else its text
            nrole = "Station"
            Dim sbase As String: sbase = Sanitize(PropStr(shp, "ID"))
            If sbase = "" Then sbase = land
            If sbase = "" Then sbase = "stop"
            nid = UniqueId(sbase, usedNode)
        ElseIf land = "start" Then        ' a node named "Start" -> id "start" (planner's origin)
            nid = UniqueId("start", usedNode)
        Else
            If land = "" Then land = "node"
            nid = UniqueId(land & NextCount(landCount, land), usedNode)
        End If
        Dim cc As Variant: cc = CenterPx(shp)
        mNodeMap.Add Array(nid, cc(0), cc(1), nrole), "k" & shp.id
        If nodeCount > 0 Then nodesJson = nodesJson & "," & vbCrLf
        nodesJson = nodesJson & NodeJson(nid, False, cc(0), cc(1), PropStr(shp, "Name"))
        nodeCount = nodeCount + 1
    Next

    For Each v In entShapes
        Set shp = v
        Dim eaid As String: eaid = ""
        If KeyExists(entOf, "k" & shp.id) Then eaid = entOf("k" & shp.id)
        Dim eid As String: eid = NodeIdForRole(shp, "Entrance", eaid, "_in", usedNode)
        Dim ec As Variant: ec = CenterPx(shp)
        mNodeMap.Add Array(eid, ec(0), ec(1), "Entrance"), "k" & shp.id
        If nodeCount > 0 Then nodesJson = nodesJson & "," & vbCrLf
        nodesJson = nodesJson & NodeJson(eid, True, ec(0), ec(1), PropStr(shp, "Name"))
        nodeCount = nodeCount + 1
        If eaid <> "" Then PutOnce assocEnt, eaid, eid, "Entrance"
    Next

    For Each v In exitShapes
        Set shp = v
        Dim xaid As String: xaid = ""
        If KeyExists(exitOf, "k" & shp.id) Then xaid = exitOf("k" & shp.id)
        Dim xid As String: xid = NodeIdForRole(shp, "Exit", xaid, "_out", usedNode)
        Dim xc As Variant: xc = CenterPx(shp)
        mNodeMap.Add Array(xid, xc(0), xc(1), "Exit"), "k" & shp.id
        If nodeCount > 0 Then nodesJson = nodesJson & "," & vbCrLf
        nodesJson = nodesJson & NodeJson(xid, True, xc(0), xc(1), PropStr(shp, "Name"))
        nodeCount = nodeCount + 1
        If xaid <> "" Then PutOnce assocExit, xaid, xid, "Exit"
    Next

    ' --- pass D: walkway edges (node ids are now final) --------------------
    For Each v In edgeLines
        Set shp = v
        Dim db As Visio.Shape, de As Visio.Shape
        Set db = Nothing: Set de = Nothing
        GetEnds shp, db, de
        Dim fromId As String, toId As String
        fromId = FinalId(db): toId = FinalId(de)
        If fromId <> "" And toId <> "" And fromId <> toId Then
            Dim ptsJson As String: ptsJson = ConnectorPointsJson(shp, fromId)
            If connCount > 0 Then connJson = connJson & "," & vbCrLf
            connJson = connJson & "  { ""from"": """ & fromId & """, ""to"": """ & toId & _
                       """, ""points"": [" & ptsJson & "] }"
            connCount = connCount + 1
        Else
            Warn "Walkway line could not resolve to two distinct nodes: " & shp.NameU
        End If
    Next

    ' --- track pass: ride animation polylines (shapes on the "Track" layer) -
    '   Each track is glued into the ride shape it belongs to. We keep its
    '   vertices in natural order (first point first) for the ride animation.
    Dim trackOf As Collection: Set trackOf = New Collection ' "k"&AttractionShapeID -> points JSON
    For Each v In trackShapes
        Set shp = v
        Dim trkAttr As Visio.Shape: Set trkAttr = TrackRideShape(shp)
        If trkAttr Is Nothing Then
            Warn "Track '" & shp.NameU & "' isn't glued to a ride - glue one end into the ride shape. Skipped."
        ElseIf CategoryOf(trkAttr) <> "ride" Then
            Warn "Track '" & shp.NameU & "' is attached to a non-ride ('" & mAttrMap("k" & trkAttr.id) & "'). Skipped."
        ElseIf KeyExists(trackOf, "k" & trkAttr.id) Then
            Warn "Ride '" & mAttrMap("k" & trkAttr.id) & "' already has a track; ignoring extra " & shp.NameU & "."
        Else
            Dim trkPts As String: trkPts = TrackPointsJson(shp)
            If trkPts = "" Then
                Warn "Track '" & shp.NameU & "' has fewer than 2 points. Skipped."
            Else
                trackOf.Add trkPts, "k" & trkAttr.id
            End If
        End If
    Next

    ' --- queue pass: per-ride wait-line polylines (shapes on the "Queue" layer)
    '   Glued into their ride like a Track. The avatar follows this from the
    '   entrance to the load point during the wait. Same order convention as a
    '   track: draw it begin (entrance side) -> end (load side).
    Dim queueOf As Collection: Set queueOf = New Collection ' "k"&AttractionShapeID -> points JSON
    For Each v In queueShapes
        Set shp = v
        Dim qAttr As Visio.Shape: Set qAttr = TrackRideShape(shp)
        If qAttr Is Nothing Then
            Warn "Queue '" & shp.NameU & "' isn't glued to a ride - glue one end into the ride shape. Skipped."
        ElseIf CategoryOf(qAttr) <> "ride" Then
            Warn "Queue '" & shp.NameU & "' is attached to a non-ride ('" & mAttrMap("k" & qAttr.id) & "'). Skipped."
        ElseIf KeyExists(queueOf, "k" & qAttr.id) Then
            Warn "Ride '" & mAttrMap("k" & qAttr.id) & "' already has a queue; ignoring extra " & shp.NameU & "."
        Else
            Dim qPts As String: qPts = TrackPointsJson(shp)
            If qPts = "" Then
                Warn "Queue '" & shp.NameU & "' has fewer than 2 points. Skipped."
            Else
                queueOf.Add qPts, "k" & qAttr.id
            End If
        End If
    Next

    ' --- pass E: attractions (entrance/exit from association lines) --------
    For Each v In attractions
        Set shp = v
        aId = mAttrMap("k" & shp.id)
        Dim ac As Variant: ac = CenterPx(shp)
        Dim entId As String, exId As String
        entId = "": exId = ""
        Dim accJson As String: accJson = ""
        If CategoryOf(shp) = "ride" Then
            If KeyExists(assocEnt, aId) Then entId = assocEnt(aId) Else _
                Warn "Ride '" & aId & "' has no Entrance link (draw a line from it to its Entrance node)."
            If KeyExists(assocExit, aId) Then exId = assocExit(aId) Else _
                Warn "Ride '" & aId & "' has no Exit link (draw a line from it to its Exit node)."
        Else
            ' restaurant/shop/pin: one or more node links (optional). entrance =
            ' exit = first; if 2+ distinct nodes, also emit accessNodeIds so the
            ' planner can route to the nearest. No link is fine (only rides need
            ' entrance/exit), so we don't warn.
            If KeyExists(directOf, "k" & shp.id) Then
                Set lst = directOf("k" & shp.id)
                ' NB: ids/cnt are proc-scoped; reset them per attraction or they
                ' accumulate every previous shop's nodes.
                Dim seenN As Collection: Set seenN = New Collection
                Dim j As Long, nid2 As String
                Dim ids As String: ids = ""
                Dim cnt As Long: cnt = 0
                For j = 1 To lst.Count
                    If KeyExists(mNodeMap, "k" & lst(j)) Then
                        nid2 = mNodeMap("k" & lst(j))(0)
                        If Not KeyExists(seenN, nid2) Then
                            seenN.Add True, nid2
                            If cnt = 0 Then entId = nid2: exId = nid2
                            If cnt > 0 Then ids = ids & ", "
                            ids = ids & """" & nid2 & """"
                            cnt = cnt + 1
                        End If
                    End If
                Next j
                If cnt >= 2 Then accJson = "[" & ids & "]"
            End If
        End If
        Dim trkJson As String: trkJson = ""
        If KeyExists(trackOf, "k" & shp.id) Then trkJson = trackOf("k" & shp.id)
        Dim queJson As String: queJson = ""
        If KeyExists(queueOf, "k" & shp.id) Then queJson = queueOf("k" & shp.id)
        If attrCount > 0 Then attrJson = attrJson & "," & vbCrLf
        attrJson = attrJson & AttractionJson(aId, ShapeName(shp), entId, exId, ac(0), ac(1), RideDur(shp), CategoryOf(shp), IsClosed(shp), WaitIdOf(shp), accJson, PropStr(shp, "Hovertext"), AvgWaitOf(shp), trkJson, queJson, LabelPosJson(shp), ThpwIdOf(shp), PropStr(shp, "Audio"), PropIsTrue(shp, Array("QInside", "QueueInside")), PropTri(shp, Array("RInside", "RideInside", "Indoor")), PropSunFrac(shp, Array("QSun", "QueueSun")))
        attrCount = attrCount + 1
    Next

    ' --- transit pass: transport lines from "transit_<Name>" layers ----------
    Dim transitCount As Long: transitCount = 0
    Dim transportJson As String: transportJson = BuildTransportJson(transitSegs, transitCount)

    ' --- geo pass: anchors from shapes carrying Prop.LatLon -------------------
    Dim geoCount As Long: geoCount = 0
    Dim geoJson As String: geoJson = BuildGeoJson(pg, geoCount)

    ' --- shelter pass: polygons on the RainCover / indoors / shade layers ------
    Dim shelterCount As Long: shelterCount = 0
    Dim sheltersJson As String: sheltersJson = BuildSheltersJson(pg, shelterCount)

    ' --- assemble blocks -----------------------------------------------------
    Dim nodesBlock As String, connBlock As String, attrBlock As String
    Dim mapBlock As String, scaleBlock As String, transportBlock As String, geoBlock As String, sheltersBlock As String
    nodesBlock = "SAMPLE.nodes = [" & vbCrLf & nodesJson & vbCrLf & "];"
    connBlock = "SAMPLE.connections = [" & vbCrLf & connJson & vbCrLf & "];"
    attrBlock = "SAMPLE.attractions = [" & vbCrLf & attrJson & vbCrLf & "];"
    mapBlock = "SAMPLE.mapExtent = " & MapExtentJson(pg) & ";"
    transportBlock = "SAMPLE.transport = [" & vbCrLf & transportJson & vbCrLf & "];"
    geoBlock = "SAMPLE.geoAnchors = [" & vbCrLf & geoJson & vbCrLf & "];"
    sheltersBlock = "SAMPLE.shelters = [" & vbCrLf & sheltersJson & vbCrLf & "];"

    Dim ftPerPx As Double: ftPerPx = ComputeScale(pg)
    If ftPerPx > 0 Then
        scaleBlock = "SAMPLE.feetPerPixel = " & JNum(ftPerPx) & ";"
    Else
        scaleBlock = "SAMPLE.feetPerPixel = null;"
    End If

    Dim outText As String   ' plain-text fallback (same blocks, copy/paste-able)
    outText = nodesBlock & vbCrLf & vbCrLf & connBlock & vbCrLf & vbCrLf & _
              attrBlock & vbCrLf & vbCrLf & mapBlock & vbCrLf & vbCrLf & scaleBlock & vbCrLf & vbCrLf & _
              transportBlock & vbCrLf & vbCrLf & geoBlock & vbCrLf & vbCrLf & sheltersBlock & vbCrLf
    Debug.Print outText

    ' --- emit: patch the park's park.js in place, else write the .txt --------
    '   (VBA And is not short-circuit, so guard with nested Ifs.)
    Dim msg As String, htmlP As String, didPatch As Boolean
    htmlP = HtmlPath()
    If htmlP <> "" Then
        If Dir$(htmlP) <> "" Then didPatch = PatchHtml(htmlP, nodesBlock, connBlock, attrBlock, mapBlock, scaleBlock, transportBlock, geoBlock, sheltersBlock)
    End If
    If didPatch Then
        msg = "Wrote " & nodeCount & " nodes, " & connCount & " connections, " & _
              attrCount & " attractions" & IIf(transitCount > 0, ", " & transitCount & " transport line(s)", "") & IIf(geoCount > 0, ", " & geoCount & " geo anchor(s)", "") & IIf(shelterCount > 0, ", " & shelterCount & " shelter polygon(s)", "") & " into:" & vbCrLf & htmlP & vbCrLf & _
              "Refresh the page in your browser."
    Else
        Dim savedTo As String: savedTo = WriteOut(outText)
        msg = "Exported " & nodeCount & " nodes, " & connCount & " connections, " & _
              attrCount & " attractions." & vbCrLf & _
              "Could not patch park.js (not found or markers missing), wrote:" & vbCrLf & _
              savedTo & vbCrLf & "Paste the blocks in manually."
    End If

    Dim svgOut As String: svgOut = ExportBackgroundSvg(mapPage, pg)
    If svgOut <> "" Then msg = msg & vbCrLf & vbCrLf & "Background page exported to:" & vbCrLf & svgOut

    If mWarnings.Count > 0 Then
        Dim wv As Variant, wAll As String
        For Each wv In mWarnings: wAll = wAll & vbCrLf & " - " & wv: Next
        Debug.Print "WARNINGS:" & wAll
        msg = msg & vbCrLf & vbCrLf & mWarnings.Count & " warning(s) - see Immediate window (Ctrl+G)."
    End If
    MsgBox msg, vbInformation, "RideSim Export"
End Sub

'--------------------------- park.js patching ---------------------------------
' Path to the park data file: HTML_PATH_OVERRIDE if set, else
' <docfolder>\parks\<active-page-slug>\park.js (one folder per park).
Private Function HtmlPath() As String
    If HTML_PATH_OVERRIDE <> "" Then HtmlPath = HTML_PATH_OVERRIDE: Exit Function
    Dim f As String: f = ParkFolder()
    If f <> "" Then HtmlPath = f & "park.js"
End Function
' <docfolder>\parks\<active-page-slug>\  — the folder holding this park's park.js
' (and background.svg). "" if the doc is unsaved.
Private Function ParkFolder() As String
    If HTML_PATH_OVERRIDE <> "" Then
        Dim k As Long: k = InStrRev(HTML_PATH_OVERRIDE, "\")
        If k > 0 Then ParkFolder = Left$(HTML_PATH_OVERRIDE, k)
        Exit Function
    End If
    Dim folder As String
    On Error Resume Next
    folder = ThisDocument.path
    On Error GoTo 0
    If folder = "" Then Exit Function
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Dim slug As String
    If Not mDataPage Is Nothing Then slug = Slugify(mDataPage.Name) Else slug = Slugify(Visio.ActivePage.Name)
    If slug = "" Then Exit Function
    ParkFolder = folder & "parks\" & slug & "\"
End Function
' The background page assigned to a page (Page.BackPage), read defensively via
' CallByName since it returns a name or a Page depending on the Visio version.
' Nothing if the page has no background page (old image-in-a-BG_LAYER model).
Private Function BackgroundPageOf(pg As Visio.Page) As Visio.Page
    On Error GoTo done
    Dim v As Variant: v = CallByName(pg, "BackPage", VbGet)
    If IsObject(v) Then
        Set BackgroundPageOf = v
    ElseIf Len(Trim$(CStr(v))) > 0 Then
        Set BackgroundPageOf = ThisDocument.Pages.Item(Trim$(CStr(v)))
    End If
done:
End Function
' The (foreground) page that uses bgPage as its background, or Nothing.
Private Function ForegroundUsing(bgPage As Visio.Page) As Visio.Page
    Dim p As Visio.Page, b As Visio.Page
    For Each p In ThisDocument.Pages
        If p.Background = 0 Then
            Set b = BackgroundPageOf(p)
            If Not b Is Nothing Then
                If b.ID = bgPage.ID Then Set ForegroundUsing = p: Exit Function
            End If
        End If
    Next p
End Function
' Export the map (background) page to background.svg in the park folder, so the
' backdrop always matches the exported data. Warns if the two pages differ in size
' (to 2 dp). Leaves the view on the data page. No-op if mapPage is Nothing.
Private Function ExportBackgroundSvg(mapPage As Visio.Page, dataPage As Visio.Page) As String
    Dim win As Visio.Window
    On Error GoTo fail
    If mapPage Is Nothing Then Exit Function
    Dim f As String: f = ParkFolder()
    If f = "" Then Exit Function
    ' warn only when the pages differ at 2 decimal places (ignore fp/precision noise)
    Dim cw As String, ch As String, bw As String, bh As String
    cw = Format$(dataPage.PageSheet.CellsU("PageWidth").ResultIU, "0.00")
    ch = Format$(dataPage.PageSheet.CellsU("PageHeight").ResultIU, "0.00")
    bw = Format$(mapPage.PageSheet.CellsU("PageWidth").ResultIU, "0.00")
    bh = Format$(mapPage.PageSheet.CellsU("PageHeight").ResultIU, "0.00")
    If cw <> bw Or ch <> bh Then _
        Warn "Map page '" & mapPage.Name & "' (" & bw & "x" & bh & " in) differs from the content page (" & cw & "x" & ch & " in); keep them the same size so nodes don't fall outside the backdrop."
    Dim svgPath As String: svgPath = f & "background.svg"
    ' For SVG, Selection.Export writes the whole page anyway (Page.Export on the
    ' object proved unreliable). MapExtentJson uses the map page's own rectangle to
    ' match. Show the map page, select all, export, then return to the data page.
    Set win = Visio.ActiveWindow
    If win Is Nothing Then Exit Function
    win.Page = mapPage.Name
    win.SelectAll
    DoEvents
    win.Selection.Export svgPath
    win.DeselectAll
    win.Page = dataPage.Name
    ExportBackgroundSvg = svgPath
    Exit Function
fail:
    On Error Resume Next
    If Not win Is Nothing Then win.Page = dataPage.Name
    Warn "Could not export background page to SVG: " & Err.Description
End Function

' Lowercase, non-alphanumerics -> single hyphen, trimmed. "Magic Kingdom" -> "magic-kingdom".
Private Function Slugify(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, lastDash As Boolean
    s = LCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch: lastDash = False
        ElseIf Not lastDash And out <> "" Then
            out = out & "-": lastDash = True
        End If
    Next
    Do While Right$(out, 1) = "-": out = Left$(out, Len(out) - 1): Loop
    Slugify = out
End Function

' Replace the text between each marker pair with its block. Returns False (and
' leaves the file untouched) if any marker is missing.
Private Function PatchHtml(path As String, nodesBlock As String, connBlock As String, _
                           attrBlock As String, mapBlock As String, scaleBlock As String, _
                           transportBlock As String, geoBlock As String, sheltersBlock As String) As Boolean
    On Error GoTo fail
    Dim s As String: s = ReadTextUtf8(path)
    Dim nl As String: nl = IIf(InStr(s, vbCrLf) > 0, vbCrLf, vbLf)
    Dim ok As Boolean: ok = True
    s = PatchSection(s, "@RIDESIM:NODES:START", "@RIDESIM:NODES:END", Reflow(nodesBlock, nl), nl, ok)
    s = PatchSection(s, "@RIDESIM:CONN:START", "@RIDESIM:CONN:END", Reflow(connBlock, nl), nl, ok)
    s = PatchSection(s, "@RIDESIM:ATTR:START", "@RIDESIM:ATTR:END", Reflow(attrBlock, nl), nl, ok)
    s = PatchSection(s, "@RIDESIM:MAP:START", "@RIDESIM:MAP:END", Reflow(mapBlock, nl), nl, ok)
    s = PatchSection(s, "@RIDESIM:SCALE:START", "@RIDESIM:SCALE:END", Reflow(scaleBlock, nl), nl, ok)
    ' transport + geo markers are newer; patch only if present so older park.js still works
    If InStr(s, "@RIDESIM:TRANSPORT:START") > 0 Then
        s = PatchSection(s, "@RIDESIM:TRANSPORT:START", "@RIDESIM:TRANSPORT:END", Reflow(transportBlock, nl), nl, ok)
    End If
    If InStr(s, "@RIDESIM:GEO:START") > 0 Then
        s = PatchSection(s, "@RIDESIM:GEO:START", "@RIDESIM:GEO:END", Reflow(geoBlock, nl), nl, ok)
    End If
    If InStr(s, "@RIDESIM:SHELTERS:START") > 0 Then
        s = PatchSection(s, "@RIDESIM:SHELTERS:START", "@RIDESIM:SHELTERS:END", Reflow(sheltersBlock, nl), nl, ok)
    End If
    If ok Then WriteTextUtf8NoBom path, s   ' only touch the file if every marker matched
    PatchHtml = ok
    Exit Function
fail:
    Warn "Could not patch park.js: " & Err.Description
    PatchHtml = False
End Function

' Keep the two marker lines; replace only the lines strictly between them.
Private Function PatchSection(s As String, startTok As String, endTok As String, _
                              block As String, nl As String, ByRef ok As Boolean) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(s, startTok): p2 = InStr(s, endTok)
    If p1 = 0 Or p2 = 0 Or p2 < p1 Then
        Warn "Marker not found in park.js: " & startTok & " .. " & endTok
        ok = False: PatchSection = s: Exit Function
    End If
    Dim eol As Long: eol = InStr(p1, s, vbLf)        ' end of the start-marker line
    If eol = 0 Then ok = False: PatchSection = s: Exit Function
    Dim lineStart As Long: lineStart = InStrRev(s, vbLf, p2) + 1  ' start of end-marker line
    PatchSection = Left$(s, eol) & block & nl & Mid$(s, lineStart)
End Function

Private Function Reflow(block As String, nl As String) As String
    Reflow = Replace(block, vbCrLf, nl)
End Function

' UTF-8 read/write (preserves emoji, em-dashes, etc.); write without BOM.
Private Function ReadTextUtf8(path As String) As String
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadTextUtf8 = st.ReadText
    st.Close
End Function

Private Sub WriteTextUtf8NoBom(path As String, ByVal text As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText text
    st.Position = 0: st.Type = 1: st.Position = 3   ' skip the 3-byte UTF-8 BOM
    Dim payload: payload = st.Read
    st.Close
    Dim bin As Object: Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1: bin.Open
    bin.Write payload
    bin.SaveToFile path, 2                           ' adSaveCreateOverWrite
    bin.Close
End Sub

'============================ helpers =========================================

Private Function MasterRole(shp As Visio.Shape) As String
    On Error Resume Next
    If shp.Master Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(shp.Master.Name)
    Dim nu As String: nu = LCase$(shp.Master.NameU)
    Select Case True
        Case nm = "attraction", nu = "attraction": MasterRole = "Attraction"
        Case nm = "restaurant", nu = "restaurant": MasterRole = "Restaurant"
        Case nm = "shop", nu = "shop", nm = "shops", nu = "shops": MasterRole = "Shop"
        Case nm = "pin", nu = "pin", nm = "pins", nu = "pins": MasterRole = "Pin"
        Case nm = "restroom", nu = "restroom", nm = "restrooms", nu = "restrooms": MasterRole = "Restroom"
        Case nm = "other", nu = "other", nm = "others", nu = "others": MasterRole = "Other"
        Case nm = "station", nu = "station", nm = "stations", nu = "stations", nm = "stop", nu = "stop": MasterRole = "Station"
        Case nm = "entrance", nu = "entrance":     MasterRole = "Entrance"
        Case nm = "exit", nu = "exit":             MasterRole = "Exit"
        Case nm = "node", nu = "node":             MasterRole = "Node"
    End Select
End Function

' Shape center in pixels (page coords -> screen-space px, Y flipped)
Private Function CenterPx(shp As Visio.Shape) As Variant
    Dim x As Double, y As Double
    x = shp.CellsU("PinX").ResultIU
    y = shp.CellsU("PinY").ResultIU
    CenterPx = Array(Round(x * PPI), Round((mPageH - y) * PPI))
End Function

' Convert a LOCAL geometry coordinate (inches) to page pixels (Y flipped)
Private Function LocalToPagePx(shp As Visio.Shape, xl As Double, yl As Double) As Variant
    Dim w As Double, h As Double, lpx As Double, lpy As Double
    Dim pinx As Double, piny As Double, ang As Double, fx As Boolean, fy As Boolean
    w = CN(shp, "Width"): h = CN(shp, "Height")
    lpx = CN(shp, "LocPinX"): lpy = CN(shp, "LocPinY")
    pinx = CN(shp, "PinX"): piny = CN(shp, "PinY")
    ang = CN(shp, "Angle"): fx = (CN(shp, "FlipX") <> 0): fy = (CN(shp, "FlipY") <> 0)
    If fx Then xl = w - xl
    If fy Then yl = h - yl
    Dim dx As Double, dy As Double, xr As Double, yr As Double
    dx = xl - lpx: dy = yl - lpy
    xr = dx * Cos(ang) - dy * Sin(ang)
    yr = dx * Sin(ang) + dy * Cos(ang)
    Dim xp As Double, yp As Double
    xp = pinx + xr: yp = piny + yr
    LocalToPagePx = Array(Round(xp * PPI), Round((mPageH - yp) * PPI))
End Function

' Optional label position (node px) = the shape's FIRST control point, used by the
' web app to place the CENTER of the attraction's name label. Skipped when there
' is no control point, or its Y Behavior (Controls col 5) is Hidden (value 6) — so
' you can leave the handle in the master but turn it off per instance.
Private Function LabelPosJson(shp As Visio.Shape) As String
    On Error GoTo none
    If Not shp.SectionExists(visSectionControls, visExistsAnywhere) Then GoTo none
    If shp.RowCount(visSectionControls) < 1 Then GoTo none
    If shp.CellsSRC(visSectionControls, 0, 5).Result(visNone) = 6 Then GoTo none   ' Y Behavior = Hidden
    Dim clx As Double, cly As Double
    clx = shp.CellsSRC(visSectionControls, 0, 0).ResultIU   ' Controls.X (local inches)
    cly = shp.CellsSRC(visSectionControls, 0, 1).ResultIU   ' Controls.Y (local inches)
    Dim lp As Variant: lp = LocalToPagePx(shp, clx, cly)
    LabelPosJson = "{ ""x"": " & CLng(lp(0)) & ", ""y"": " & CLng(lp(1)) & " }"
    Exit Function
none:
    LabelPosJson = ""
End Function

' Where the app stretches background.svg, in node px. Prefers the new background-
' page model; falls back to a BG_LAYER ("Bck") shape. "null" if neither.
Private Function MapExtentJson(pg As Visio.Page) As String
    ' Preferred (new model): the map is on the background page and its SVG export is
    ' the WHOLE page. The extent is the map page rectangle — its own width/height (so
    ' the SVG scales 1:1) shifted down by any page-height difference (the pages overlay
    ' at the bottom-left origin, but node Y is measured from the DATA page's top).
    ' Identical page sizes -> { 0, 0, pageW, pageH }. A stray "Bck" shape is ignored.
    Dim mp As Visio.Page: Set mp = BackgroundPageOf(pg)
    If Not mp Is Nothing Then
        Dim mpw As Double, mph As Double
        mpw = mp.PageSheet.CellsU("PageWidth").ResultIU
        mph = mp.PageSheet.CellsU("PageHeight").ResultIU
        MapExtentJson = "{ ""x"": 0, ""y"": " & CLng((mPageH - mph) * PPI) & _
            ", ""w"": " & CLng(mpw * PPI) & ", ""h"": " & CLng(mph * PPI) & " }"
        Exit Function
    End If
    ' Legacy: bounding box of a shape on the BG_LAYER layer of THIS page (you supply
    ' background.png/.svg yourself).
    Dim shp As Visio.Shape, found As Boolean
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    minX = 1E+30: minY = 1E+30: maxX = -1E+30: maxY = -1E+30
    For Each shp In pg.Shapes
        ' only the image itself - never nodes/attractions/scale shapes on the layer
        If OnLayer(shp, BG_LAYER) And MasterRole(shp) = "" And ScaleRole(shp) = "" Then
            found = True
            Dim w As Double, h As Double, i As Long, pt As Variant
            w = CN(shp, "Width"): h = CN(shp, "Height")
            Dim corners(0 To 3) As Variant
            corners(0) = LocalToPagePx(shp, 0, 0)
            corners(1) = LocalToPagePx(shp, w, 0)
            corners(2) = LocalToPagePx(shp, w, h)
            corners(3) = LocalToPagePx(shp, 0, h)
            For i = 0 To 3
                pt = corners(i)
                If pt(0) < minX Then minX = pt(0)
                If pt(0) > maxX Then maxX = pt(0)
                If pt(1) < minY Then minY = pt(1)
                If pt(1) > maxY Then maxY = pt(1)
            Next i
        End If
    Next shp
    If found Then
        MapExtentJson = "{ ""x"": " & CLng(minX) & ", ""y"": " & CLng(minY) & _
            ", ""w"": " & CLng(maxX - minX) & ", ""h"": " & CLng(maxY - minY) & " }"
        Exit Function
    End If
    Warn "No background page and no BG_LAYER shape - map extent not exported (background won't auto-align)."
    MapExtentJson = "null"
End Function

' True if a shape belongs to a layer with the given name.
Private Function OnLayer(shp As Visio.Shape, layerName As String) As Boolean
    On Error Resume Next
    Dim i As Long
    For i = 1 To shp.LayerCount
        If LCase$(shp.Layer(i).Name) = LCase$(layerName) Then OnLayer = True: Exit Function
    Next i
End Function

' Transport lines live on the "transit_<LineName>" layers (prefix at module top).
Private Function IsTransitLayer(shp As Visio.Shape) As Boolean
    On Error Resume Next
    Dim i As Long
    For i = 1 To shp.LayerCount
        If LCase$(Left$(shp.Layer(i).Name, Len(TRANSIT_PREFIX))) = TRANSIT_PREFIX Then IsTransitLayer = True: Exit Function
    Next i
End Function
' The line name = the layer name after the "transit_" prefix.
Private Function TransitLineOf(shp As Visio.Shape) As String
    On Error Resume Next
    Dim i As Long, nm As String
    For i = 1 To shp.LayerCount
        nm = shp.Layer(i).Name
        If LCase$(Left$(nm, Len(TRANSIT_PREFIX))) = TRANSIT_PREFIX Then TransitLineOf = Mid$(nm, Len(TRANSIT_PREFIX) + 1): Exit Function
    Next i
End Function

' Build SAMPLE.transport JSON from the directed transit segments. Segments are
' grouped by their transit_<Name> layer; each segment becomes a directed
' {from, to, minutes, points}; stops are derived from the segment endpoints,
' carrying any ThPWID / AvgWait from the Station shape (the boarding wait).
Private Function BuildTransportJson(segs As Collection, ByRef lineCount As Long) As String
    lineCount = 0
    If segs Is Nothing Then Exit Function
    If segs.Count = 0 Then Exit Function
    Dim names As Collection: Set names = New Collection   ' distinct line names, in order
    Dim v As Variant, shp As Visio.Shape, ln As String
    For Each v In segs
        Set shp = v: ln = TransitLineOf(shp)
        If ln <> "" And Not KeyExists(names, ln) Then names.Add ln, ln
    Next
    Dim out As String: out = ""
    Dim nameV As Variant
    For Each nameV In names
        ln = CStr(nameV)
        Dim stopIds As Collection: Set stopIds = New Collection      ' nodeId -> True (dedup)
        Dim stopList As Collection: Set stopList = New Collection     ' Array(nodeId, stationShape)
        Dim segJson As String: segJson = "": Dim segFirst As Boolean: segFirst = True
        For Each v In segs
            Set shp = v
            If TransitLineOf(shp) = ln Then
                ' direction comes from the FromLink / ToLink cell references (Visio
                ' won't glue a 2-D station, so we point cells at the stops instead)
                Dim b As Visio.Shape, e As Visio.Shape
                Set b = LinkedStation(shp, True)
                Set e = LinkedStation(shp, False)
                If b Is Nothing Or e Is Nothing Then
                    Warn "Transit segment on '" & ln & "' is missing a FromLink/ToLink reference to a Station: " & shp.NameU
                Else
                    Dim fromId As String, toId As String
                    fromId = FinalId(b): toId = FinalId(e)
                    If fromId = "" Or toId = "" Or fromId = toId Then
                        Warn "Transit segment on '" & ln & "' could not resolve two distinct Stops: " & shp.NameU
                    Else
                        Dim mins As Double: mins = RideDur(shp)
                        Dim rev As Boolean: rev = PropBool(shp, "Reverse")  ' copy a segment + flip for the other direction
                        Dim pts As String: pts = SegmentPointsJson(shp, rev)
                        If Not segFirst Then segJson = segJson & "," & vbCrLf
                        segJson = segJson & "    { ""from"": """ & fromId & """, ""to"": """ & toId & _
                                  """, ""minutes"": " & JNum(mins) & ", ""path"": [" & pts & "] }"
                        segFirst = False
                        If Not KeyExists(stopIds, fromId) Then stopIds.Add True, fromId: stopList.Add Array(fromId, b)
                        If Not KeyExists(stopIds, toId) Then stopIds.Add True, toId: stopList.Add Array(toId, e)
                    End If
                End If
            End If
        Next
        If Not segFirst Then            ' line produced at least one usable segment
            Dim stopsJson As String: stopsJson = "": Dim sFirst As Boolean: sFirst = True
            Dim sv As Variant
            For Each sv In stopList
                Dim arr As Variant: arr = sv
                Dim sid As String: sid = arr(0)
                Dim sShp As Visio.Shape: Set sShp = arr(1)
                Dim obj As String: obj = "{ ""node"": """ & sid & """"
                If Not sShp Is Nothing Then
                    Dim tw As String: tw = ThpwIdOf(sShp)
                    If Trim$(tw) <> "" Then obj = obj & ", ""thpwId"": """ & JStr(Trim$(tw)) & """"
                    Dim aw As Double: aw = AvgWaitOf(sShp)
                    If aw >= 0 Then obj = obj & ", ""avgWait"": " & JNum(aw)
                End If
                obj = obj & " }"
                If Not sFirst Then stopsJson = stopsJson & ", "
                stopsJson = stopsJson & obj: sFirst = False
            Next
            If lineCount > 0 Then out = out & "," & vbCrLf
            out = out & "  { ""id"": """ & Slugify(ln) & """, ""name"": """ & JStr(ln) & """," & vbCrLf & _
                  "    ""stops"": [" & stopsJson & "]," & vbCrLf & _
                  "    ""segments"": [" & vbCrLf & segJson & vbCrLf & "    ] }"
            lineCount = lineCount + 1
        End If
    Next
    BuildTransportJson = out
End Function

' Build SAMPLE.geoAnchors from shapes carrying Prop.LatLon ("lat,lon"). Each
' becomes { x, y, lat, lon } using the shape's center pixel; the web app fits an
' affine GPS->pixel transform from 3+ of them.
Private Function BuildGeoJson(pg As Visio.Page, ByRef cnt As Long) As String
    cnt = 0
    Dim out As String, shp As Visio.Shape
    For Each shp In pg.Shapes
        ' LatLon is an optional field that sits blank on most master instances, so
        ' anything that isn't a real "lat,lon" pair is skipped silently (no spam);
        ' only a genuine pair that's out of range is worth a warning.
        Dim raw As String: raw = Trim$(LatLonOf(shp))
        If InStr(raw, ",") > 0 Then
            Dim parts() As String: parts = Split(raw, ",")
            Dim lat As Double, lon As Double
            lat = ParseSigned(parts(0)): lon = ParseSigned(parts(1))
            If Abs(lat) <= 90 And Abs(lon) <= 180 And (lat <> 0 Or lon <> 0) Then
                Dim c As Variant: c = CenterPx(shp)
                If cnt > 0 Then out = out & "," & vbCrLf
                out = out & "  { ""x"": " & CLng(c(0)) & ", ""y"": " & CLng(c(1)) & _
                      ", ""lat"": " & JGeo(lat) & ", ""lon"": " & JGeo(lon) & " }"
                cnt = cnt + 1
            Else
                Warn "Prop.LatLon out of range on '" & shp.NameU & "': " & raw
            End If
        End If
    Next shp
    BuildGeoJson = out
End Function

' Build SAMPLE.shelters from polygons on the RainCover / indoors / shade layers.
' A shape can be on more than one layer, so each emits the flags it carries:
'   cover  (RainCover) -> keeps rain off,  indoor (indoors) -> a building you can
'   duck into,  shade (shade) -> shaded but maybe not covered.
Private Function BuildSheltersJson(pg As Visio.Page, ByRef cnt As Long) As String
    cnt = 0
    Dim out As String, shp As Visio.Shape
    For Each shp In pg.Shapes
        Dim cover As Boolean, indoor As Boolean, shade As Boolean
        cover = OnLayer(shp, "RainCover")
        indoor = OnLayer(shp, "indoors")
        shade = OnLayer(shp, "shade")
        If cover Or indoor Or shade Then
            Dim pj As String: pj = PolygonPointsJson(shp)
            If pj <> "" Then
                If cnt > 0 Then out = out & "," & vbCrLf
                out = out & "  { ""points"": [" & pj & "]"
                If cover Then out = out & ", ""cover"": true"
                If indoor Then out = out & ", ""indoor"": true"
                If shade Then out = out & ", ""shade"": true"
                out = out & " }"
                cnt = cnt + 1
            End If
        End If
    Next shp
    BuildSheltersJson = out
End Function

' A shape's outline as JSON point objects (inner list, no brackets), in page px.
' Uses Shape.Paths, which returns the flattened outline already transformed into
' the parent (page) coordinate system — Visio applies pin/angle/flip/size for us.
' That fixes primitives like the Rectangle tool, whose Geometry cells are Width/
' Height *formulas* that the raw-cell reader placed wrong. Empty for <3 points.
Private Function PolygonPointsJson(shp As Visio.Shape) As String
    Dim pts As Collection: Set pts = New Collection
    On Error Resume Next
    Dim pth As Visio.Path, arr() As Double, k As Long
    For Each pth In shp.Paths
        Call pth.Points(0.05, arr)             ' flatness (inches); small = smoother curves; fills arr() ByRef
        For k = LBound(arr) To UBound(arr) - 1 Step 2
            pts.Add Array(Round(arr(k) * PPI), Round((mPageH - arr(k + 1)) * PPI))
        Next k
    Next pth
    On Error GoTo 0
    If pts.Count < 3 Then Exit Function
    Dim s As String, kk As Long, pv As Variant
    For kk = 1 To pts.Count
        pv = pts(kk)
        If kk > 1 Then s = s & ", "
        s = s & "{ ""x"": " & CLng(pv(0)) & ", ""y"": " & CLng(pv(1)) & " }"
    Next kk
    PolygonPointsJson = s
End Function

' "lat,lon" from Shape Data. Tries the internal name first, then the visible
' LABEL (a typed LatLon field usually gets an auto internal name like Prop.Row_1).
Private Function LatLonOf(shp As Visio.Shape) As String
    On Error Resume Next
    Dim names As Variant: names = Array("LatLon", "LatLong", "LatLng", "GPS")
    Dim i As Long, Row As Long, vstr As String
    For i = LBound(names) To UBound(names)
        If shp.CellExistsU("Prop." & names(i), 0) Then
            vstr = Trim$(shp.CellsU("Prop." & names(i)).ResultStr(""))
            If vstr <> "" Then LatLonOf = vstr: Exit Function
        End If
    Next i
    For Row = 0 To shp.RowCount(visSectionProp) - 1
        Dim lbl As String: lbl = shp.CellsSRC(visSectionProp, Row, 2).ResultStr(visNone)
        For i = LBound(names) To UBound(names)
            If lbl Like names(i) Then
                vstr = Trim$(shp.CellsSRC(visSectionProp, Row, 0).ResultStr(""))
                If vstr <> "" Then LatLonOf = vstr: Exit Function
            End If
        Next i
    Next Row
End Function

' Signed decimal parse (ParseNum drops the sign, which breaks W lon / S lat).
Private Function ParseSigned(ByVal s As String) As Double
    s = Trim$(s)
    Dim neg As Boolean: neg = (Left$(s, 1) = "-")
    Dim v As Double: v = ParseNum(s)
    ParseSigned = IIf(neg, -v, v)
End Function

' Higher-precision JSON number for lat/lon (7 decimals ~= 1cm).
Private Function JGeo(d As Double) As String
    JGeo = Replace(Format$(d, "0.0000000"), ",", ".")
End Function

' The Station a transit segment links to via its FromLink (isFrom=True) or ToLink
' cell. Visio won't glue a 2-D station, so the segment instead carries a cell
' whose FORMULA references the station shape; we read that reference. Looks in a
' few likely cell homes and returns the referenced (known) shape.
Private Function LinkedStation(seg As Visio.Shape, isFrom As Boolean) As Visio.Shape
    On Error Resume Next
    Dim base As String: base = IIf(isFrom, "FromLink", "ToLink")
    Dim cands(0 To 4) As String
    cands(0) = "Connections." & base & ".X"
    cands(1) = "Connections." & base & ".Y"
    cands(2) = "User." & base
    cands(3) = "Prop." & base
    cands(4) = "Scratch." & base       ' unlikely, but harmless
    Dim i As Long
    For i = 0 To 4
        If seg.CellExistsU(cands(i), 0) Then
            Dim sh As Visio.Shape: Set sh = RefShape(seg.ContainingPage, seg.CellsU(cands(i)).FormulaU)
            If Not sh Is Nothing Then Set LinkedStation = ClimbKnown(sh): Exit Function
        End If
    Next i
End Function

' Resolve the shape a ShapeSheet formula references: the identifier just before
' the first "!" (handles 'Quoted Name'!.. and Sheet.N!.. forms).
Private Function RefShape(pg As Visio.Page, formula As String) As Visio.Shape
    On Error Resume Next
    Dim bang As Long: bang = InStr(formula, "!")
    If bang <= 1 Then Exit Function
    Dim tok As String, j As Long: j = bang - 1
    If Mid$(formula, j, 1) = "'" Then                       ' 'Some Name'!
        Dim q1 As Long: q1 = InStrRev(formula, "'", j - 1)
        If q1 > 0 Then tok = Mid$(formula, q1 + 1, j - q1 - 1)
    Else                                                   ' Sheet.7! / Name!
        Dim p As Long: p = j
        Do While p >= 1
            Dim ch As String: ch = Mid$(formula, p, 1)
            If (ch Like "[A-Za-z0-9]") Or ch = "." Or ch = "_" Then p = p - 1 Else Exit Do
        Loop
        tok = Mid$(formula, p + 1, j - p)
    End If
    If tok = "" Then Exit Function
    Dim sh As Visio.Shape: Set sh = pg.Shapes.ItemU(tok)   ' NameU (e.g. Sheet.7)
    If sh Is Nothing Then                                  ' fall back to display name
        Dim t As Visio.Shape
        For Each t In pg.Shapes
            If LCase$(t.Name) = LCase$(tok) Or LCase$(t.NameU) = LCase$(tok) Then Set sh = t: Exit For
        Next t
    End If
    Set RefShape = sh
End Function

' A transit segment's geometry as JSON point objects (inner list, no brackets).
' reverseIt flips the order - lets you copy a one-way segment for the return
' direction and set Prop.Reverse=TRUE instead of redrawing it backwards.
Private Function SegmentPointsJson(shp As Visio.Shape, reverseIt As Boolean) As String
    Dim pts As Collection: Set pts = New Collection
    On Error Resume Next
    Dim i As Long, r As Long, xl As Double, yl As Double
    For i = 1 To shp.GeometryCount
        r = 1
        Do While shp.CellExistsU("Geometry" & i & ".X" & r, 0)
            xl = shp.CellsU("Geometry" & i & ".X" & r).ResultIU
            yl = shp.CellsU("Geometry" & i & ".Y" & r).ResultIU
            pts.Add LocalToPagePx(shp, xl, yl)
            r = r + 1
        Loop
    Next i
    If pts.Count < 2 Then
        Set pts = New Collection
        pts.Add Array(Round(CN(shp, "BeginX") * PPI), Round((mPageH - CN(shp, "BeginY")) * PPI))
        pts.Add Array(Round(CN(shp, "EndX") * PPI), Round((mPageH - CN(shp, "EndY")) * PPI))
    End If
    On Error GoTo 0
    If reverseIt Then Set pts = ReverseCol(pts)
    Dim s As String, kk As Long, pv As Variant
    For kk = 1 To pts.Count
        pv = pts(kk)
        If kk > 1 Then s = s & ", "
        s = s & "{ ""x"": " & CLng(pv(0)) & ", ""y"": " & CLng(pv(1)) & " }"
    Next kk
    SegmentPointsJson = s
End Function

' Read a boolean Shape Data / cell value (TRUE / 1 / yes).
Private Function PropBool(shp As Visio.Shape, propName As String) As Boolean
    On Error Resume Next
    Dim s As String: s = LCase$(Trim$(PropStr(shp, propName)))
    If s = "true" Or s = "1" Or s = "yes" Then PropBool = True: Exit Function
    If shp.CellExistsU("Prop." & propName, 0) Then
        If shp.CellsU("Prop." & propName).Result(visNone) <> 0 Then PropBool = True
    End If
End Function

' "ScaleStart" / "ScaleEnd" / "" - by master name or shape name.
Private Function ScaleRole(shp As Visio.Shape) As String
    On Error Resume Next
    Dim nm As String
    If Not shp.Master Is Nothing Then nm = LCase$(shp.Master.Name)
    Dim nu As String: nu = LCase$(shp.NameU)
    If nm = "scalestart" Or InStr(nu, "scalestart") > 0 Then ScaleRole = "ScaleStart": Exit Function
    If nm = "scaleend" Or InStr(nu, "scaleend") > 0 Then ScaleRole = "ScaleEnd"
End Function

' True if a line's two endpoints are the scale-bar shapes (either direction).
Private Function IsScaleLine(b As Visio.Shape, e As Visio.Shape) As Boolean
    If b Is Nothing Or e Is Nothing Then Exit Function
    IsScaleLine = KeyExists(mScaleIds, "k" & b.id) And KeyExists(mScaleIds, "k" & e.id)
End Function

' Feet-per-pixel from the scale bar: feet (text on the connecting line, else on
' a scale shape) divided by the pixel distance between ScaleStart and ScaleEnd.
Private Function ComputeScale(pg As Visio.Page) As Double
    Dim shp As Visio.Shape, sShp As Visio.Shape, eShp As Visio.Shape
    For Each shp In pg.Shapes
        Select Case ScaleRole(shp)
            Case "ScaleStart": Set sShp = shp
            Case "ScaleEnd": Set eShp = shp
        End Select
    Next
    If sShp Is Nothing Or eShp Is Nothing Then Exit Function   ' no scale bar -> 0

    Dim feet As Double, ln As Visio.Shape
    For Each shp In pg.Shapes                                   ' find the connecting line
        If shp.OneD And MasterRole(shp) = "" Then
            Dim b As Visio.Shape, e As Visio.Shape
            Set b = Nothing: Set e = Nothing
            GetEnds shp, b, e
            If Not b Is Nothing And Not e Is Nothing Then
                If (b.id = sShp.id And e.id = eShp.id) Or (b.id = eShp.id And e.id = sShp.id) Then
                    Set ln = shp: Exit For
                End If
            End If
        End If
    Next
    If Not ln Is Nothing Then feet = ParseNum(ShapeText(ln))
    If feet <= 0 Then feet = ParseNum(ShapeText(sShp))
    If feet <= 0 Then feet = ParseNum(ShapeText(eShp))
    If feet <= 0 Then
        Warn "Found ScaleStart/ScaleEnd but no distance text (feet) - scale not set."
        Exit Function
    End If

    Dim cs As Variant, ce As Variant
    cs = CenterPx(sShp): ce = CenterPx(eShp)
    Dim d As Double: d = Sqr((cs(0) - ce(0)) ^ 2 + (cs(1) - ce(1)) ^ 2)
    If d <= 0 Then Exit Function
    ComputeScale = feet / d
End Function

' First number in a string ("350 ft" -> 350, "12.5" -> 12.5). 0 if none.
Private Function ParseNum(ByVal s As String) As Double
    Dim i As Long, ch As String, num As String, seenDot As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            num = num & ch
        ElseIf ch = "." And Not seenDot And num <> "" Then
            num = num & ".": seenDot = True
        ElseIf num <> "" Then
            Exit For
        End If
    Next i
    If num <> "" Then ParseNum = val(num)   ' Val uses "." regardless of locale
End Function

' Number -> locale-safe JSON literal (always "." decimal, no scientific).
Private Function JNum(d As Double) As String
    JNum = Replace(Format$(d, "0.######"), ",", ".")
End Function

' Build the JSON points list for a connector, ordered begin(fromId) -> end.
Private Function ConnectorPointsJson(shp As Visio.Shape, fromId As String) As String
    Dim pts As Collection: Set pts = New Collection
    On Error Resume Next

    ' walk every geometry section's vertices (approximates arcs by endpoints)
    Dim i As Long, r As Long
    For i = 1 To shp.GeometryCount
        r = 1
        Do While shp.CellExistsU("Geometry" & i & ".X" & r, 0)
            Dim xl As Double, yl As Double
            xl = shp.CellsU("Geometry" & i & ".X" & r).ResultIU
            yl = shp.CellsU("Geometry" & i & ".Y" & r).ResultIU
            pts.Add LocalToPagePx(shp, xl, yl)
            r = r + 1
        Loop
    Next i

    ' fallback: straight begin/end in page coords
    If pts.Count < 2 Then
        Set pts = New Collection
        pts.Add Array(Round(CN(shp, "BeginX") * PPI), Round((mPageH - CN(shp, "BeginY")) * PPI))
        pts.Add Array(Round(CN(shp, "EndX") * PPI), Round((mPageH - CN(shp, "EndY")) * PPI))
    End If
    On Error GoTo 0

    ' orient so the first point is nearest the "from" node
    Dim fromC As Variant: fromC = NodeCenter(fromId)
    If Not IsEmpty(fromC) And pts.Count >= 2 Then
        Dim p1 As Variant, pN As Variant
        p1 = pts(1): pN = pts(pts.Count)
        If Dist2(pN, fromC) < Dist2(p1, fromC) Then Set pts = ReverseCol(pts)
    End If

    Dim s As String, k As Long, pv As Variant
    For k = 1 To pts.Count
        pv = pts(k)
        If k > 1 Then s = s & ", "
        s = s & "{ ""x"": " & CLng(pv(0)) & ", ""y"": " & CLng(pv(1)) & " }"
    Next k
    ConnectorPointsJson = s
End Function

' The ride Attraction a Track shape is glued into (via any of its connections).
Private Function TrackRideShape(trk As Visio.Shape) As Visio.Shape
    On Error Resume Next
    Dim cx As Visio.Connect, k As Visio.Shape
    For Each cx In trk.Connects               ' connections FROM the track TO other shapes
        Set k = ClimbKnown(cx.ToSheet)
        If Not k Is Nothing Then
            If RoleOfShape(k) = "Attraction" Then Set TrackRideShape = k: Exit Function
        End If
    Next cx
End Function

' Ordered vertices of a Track shape as a JSON points array (natural order, first
' vertex first). "" if the shape has fewer than 2 points.
Private Function TrackPointsJson(shp As Visio.Shape) As String
    Dim pts As Collection: Set pts = New Collection
    On Error Resume Next
    ' Shape.Paths returns the flattened outline in page coordinates, so curves
    ' (arcs, splines) come through as points automatically — no need to hand-add
    ' vertices in the drawing to make a curve export smoothly.
    Dim pth As Visio.Path, arr() As Double, k As Long
    For Each pth In shp.Paths
        Call pth.Points(TRACK_FLATNESS, arr)   ' smaller flatness = smoother curves; fills arr() ByRef
        For k = LBound(arr) To UBound(arr) - 1 Step 2
            pts.Add Array(Round(arr(k) * PPI), Round((mPageH - arr(k + 1)) * PPI))
        Next k
    Next pth
    If pts.Count < 2 Then                       ' straight connector with no path: its two endpoints
        Set pts = New Collection
        pts.Add Array(Round(CN(shp, "BeginX") * PPI), Round((mPageH - CN(shp, "BeginY")) * PPI))
        pts.Add Array(Round(CN(shp, "EndX") * PPI), Round((mPageH - CN(shp, "EndY")) * PPI))
    End If
    On Error GoTo 0
    If pts.Count < 2 Then Exit Function
    Dim s As String, kk As Long, pv As Variant
    For kk = 1 To pts.Count
        pv = pts(kk)
        If kk > 1 Then s = s & ", "
        s = s & "{ ""x"": " & CLng(pv(0)) & ", ""y"": " & CLng(pv(1)) & " }"
    Next kk
    TrackPointsJson = "[" & s & "]"
End Function

' Determine the begin/end node shapes a connector is glued to.
Private Sub GetEnds(CN As Visio.Shape, ByRef bShp As Visio.Shape, ByRef eShp As Visio.Shape)
    On Error Resume Next
    Dim cx As Visio.Connect
    For Each cx In CN.Connects
        Dim nm As String: nm = cx.FromCell.Name
        If InStr(1, nm, "Begin", vbTextCompare) > 0 Then
            Set bShp = cx.ToSheet
        ElseIf InStr(1, nm, "End", vbTextCompare) > 0 Then
            Set eShp = cx.ToSheet
        End If
    Next cx
End Sub

' Climb from a (possibly sub-) shape to the nearest enclosing shape that is one
' of our known shapes (a graph node or an attraction). Nothing if none.
Private Function ClimbKnown(shp As Visio.Shape) As Visio.Shape
    On Error Resume Next
    Dim s As Visio.Shape: Set s = shp
    Do While Not s Is Nothing
        If KeyExists(mRole, "k" & s.id) Or KeyExists(mAttrMap, "k" & s.id) Then
            Set ClimbKnown = s: Exit Function
        End If
        If s.ContainingShape Is Nothing Then Exit Do
        If s.ContainingShape.id = s.id Then Exit Do
        Set s = s.ContainingShape
    Loop
End Function

Private Function RoleOfShape(s As Visio.Shape) As String
    If KeyExists(mAttrMap, "k" & s.id) Then RoleOfShape = "Attraction": Exit Function
    If KeyExists(mRole, "k" & s.id) Then RoleOfShape = mRole("k" & s.id)
End Function

' Web category from the shape's master role. Attractions are all stored
' together, so the original master role is kept in mRole and read back here.
Private Function CategoryOf(shp As Visio.Shape) As String
    CategoryOf = "ride"
    If KeyExists(mRole, "k" & shp.id) Then
        Select Case mRole("k" & shp.id)
            Case "Restaurant": CategoryOf = "restaurant"
            Case "Shop":       CategoryOf = "shop"
            Case "Pin":        CategoryOf = "pin"
            Case "Restroom":   CategoryOf = "restroom"
            Case "Other":      CategoryOf = "other"
        End Select
    End If
End Function

' Final node id for a (possibly sub-) shape, after pass C assigned ids.
Private Function FinalId(shp As Visio.Shape) As String
    Dim s As Visio.Shape: Set s = ClimbKnown(shp)
    If s Is Nothing Then Exit Function
    If KeyExists(mNodeMap, "k" & s.id) Then FinalId = mNodeMap("k" & s.id)(0)
End Function

' Record an attraction->entrance/exit nodeId association; warn on a second.
Private Sub PutOnce(c As Collection, key As String, val As String, label As String)
    If KeyExists(c, key) Then
        Warn "Attraction '" & key & "' has multiple " & label & " links; keeping '" & _
             c(key) & "', ignoring '" & val & "'."
    Else
        c.Add val, key
    End If
End Sub

' Record which attraction an Entrance/Exit shape belongs to; warn on a second.
Private Sub PutOnceKey(c As Collection, key As String, val As String, label As String, lineShp As Visio.Shape)
    If KeyExists(c, key) Then
        Warn label & " node already linked to attraction '" & c(key) & _
             "'; ignoring extra link from line " & lineShp.NameU & "."
    Else
        c.Add val, key
    End If
End Sub

'--------------------------- id / value readers -------------------------------
' Attraction id: Prop.ID, else sanitized shape text, else "a"&ShapeID. Unique.
Private Function AttrIdFor(shp As Visio.Shape, used As Collection) As String
    Dim p As String: p = PropStr(shp, "ID")
    If p <> "" Then AttrIdFor = UniqueId(p, used): Exit Function
    Dim t As String: t = Sanitize(ShapeText(shp))
    If t <> "" Then AttrIdFor = UniqueId(t, used): Exit Function
    AttrIdFor = UniqueId("a" & shp.id, used)
End Function

' Entrance/Exit id: Prop.ID, else <attractionId><suffix> (e.g. pirates_in).
Private Function NodeIdForRole(shp As Visio.Shape, role As String, attrId As String, _
                               suffix As String, used As Collection) As String
    Dim p As String: p = PropStr(shp, "ID")
    If p <> "" Then NodeIdForRole = UniqueId(p, used): Exit Function
    If attrId <> "" Then NodeIdForRole = UniqueId(attrId & suffix, used): Exit Function
    Warn role & " node (Shape ID " & shp.id & ") is not linked to an Attraction - " & _
         "draw a line from its Attraction shape to it."
    NodeIdForRole = UniqueId(IIf(role = "Entrance", "ent", "ex") & shp.id, used)
End Function

' Lowercase, keep a-z0-9, turn spaces/-/_ into single underscores, trim them.
Private Function Sanitize(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = LCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
        ElseIf ch = " " Or ch = "-" Or ch = "_" Then
            out = out & "_"
        End If
    Next i
    Do While InStr(out, "__") > 0: out = Replace(out, "__", "_"): Loop
    Do While Len(out) > 0 And Left$(out, 1) = "_": out = Mid$(out, 2): Loop
    Do While Len(out) > 0 And Right$(out, 1) = "_": out = Left$(out, Len(out) - 1): Loop
    Sanitize = out
End Function

' Return base, or base_2, base_3... so it is unique; records it in `used`.
Private Function UniqueId(ByVal base As String, used As Collection) As String
    If base = "" Then base = "x"
    Dim id As String: id = base
    Dim n As Long: n = 2
    Do While KeyExists(used, id)
        id = base & "_" & n: n = n + 1
    Loop
    used.Add id, id
    UniqueId = id
End Function

' Per-key running counter (1,2,3...) stored in a Collection.
Private Function NextCount(c As Collection, key As String) As Long
    Dim n As Long
    If KeyExists(c, key) Then n = c(key): c.Remove key
    n = n + 1
    c.Add n, key
    NextCount = n
End Function

Private Function ShapeText(shp As Visio.Shape) As String
    On Error Resume Next
    ShapeText = Trim$(shp.text)
End Function

' Order node shapes left->right, then bottom->top (for stable auto numbering).
Private Function SortNodesSpatially(coll As Collection) As Collection
    Dim res As New Collection
    Dim nN As Long: nN = coll.Count
    If nN = 0 Then Set SortNodesSpatially = res: Exit Function
    Dim sh() As Visio.Shape, xs() As Double, ys() As Double, idx() As Long
    ReDim sh(1 To nN): ReDim xs(1 To nN): ReDim ys(1 To nN): ReDim idx(1 To nN)
    Dim i As Long
    For i = 1 To nN
        Set sh(i) = coll(i)
        Dim cc As Variant: cc = CenterPx(sh(i))
        xs(i) = cc(0): ys(i) = cc(1): idx(i) = i
    Next i
    Dim j As Long, k As Long, t As Long
    For j = 2 To nN                     ' insertion sort on idx
        k = j
        Do While k > 1
            If LessNode(xs(idx(k)), ys(idx(k)), xs(idx(k - 1)), ys(idx(k - 1))) Then
                t = idx(k): idx(k) = idx(k - 1): idx(k - 1) = t
                k = k - 1
            Else
                Exit Do
            End If
        Loop
    Next j
    For i = 1 To nN: res.Add sh(idx(i)): Next i
    Set SortNodesSpatially = res
End Function

' True if (ax,ay) sorts before (bx,by): left->right, ties (same column within
' EPS px) bottom->top (larger y first, since y grows downward on screen).
Private Function LessNode(ax As Double, ay As Double, bx As Double, by As Double) As Boolean
    Const EPS As Double = 8
    If Abs(ax - bx) > EPS Then LessNode = (ax < bx) Else LessNode = (ay > by)
End Function

' Ride duration (minutes) from the attraction's Shape Data. Tries common field
' names; numeric or text ("12 min") both work. Falls back to DEFAULT_RIDE.
Private Function RideDur(shp As Visio.Shape) As Double
    On Error Resume Next
    Dim names As Variant: names = Array("RideDuration", "Duration", "RideTime", "Ride", "Minutes")
    Dim i As Long, cell As String, v As Double
    Dim Row As Long
    For Row = 0 To shp.RowCount(visSectionProp) - 1
        For i = LBound(names) To UBound(names)
            cell = "Prop." & names(i)
            If shp.CellsSRC(visSectionProp, Row, 2).ResultStr(visNone) Like names(i) Then
                v = shp.CellsSRC(visSectionProp, Row, 0).Result(visNone)                 ' numeric shape data
                If v <= 0 Then v = ParseNum(shp.CellsU(cell).ResultStr(""))  ' text shape data
                If v > 0 Then RideDur = v: Exit Function
            End If
        Next i
    Next Row
    RideDur = DEFAULT_RIDE
End Function

Private Function ShapeName(shp As Visio.Shape) As String
    Dim p As String: p = PropStr(shp, "Name")
    If p <> "" Then ShapeName = p: Exit Function
    On Error Resume Next
    ShapeName = Trim$(shp.text)
End Function

' Average wait (minutes) from Shape Data, if present. Returns -1 when unset so
' the exporter only emits "avgWait" for rides that actually carry it.
Private Function AvgWaitOf(shp As Visio.Shape) As Double
    On Error Resume Next
    AvgWaitOf = -1
    Dim names As Variant: names = Array("AvgWait", "AverageWait", "AvgWaitTime", "WaitMinutes")
    Dim i As Long, cell As String, Row As Long
    For Row = 0 To shp.RowCount(visSectionProp) - 1
        For i = LBound(names) To UBound(names)
            cell = "Prop." & names(i)
            If shp.CellsSRC(visSectionProp, Row, 2).ResultStr(visNone) Like names(i) Then
                Dim raw As String: raw = Trim$(shp.CellsSRC(visSectionProp, Row, 0).ResultStr(""))
                If raw <> "" Then
                    Dim v As Double: v = shp.CellsSRC(visSectionProp, Row, 0).Result(visNone)
                    If v <= 0 Then v = ParseNum(raw)
                    If v >= 0 Then AvgWaitOf = v: Exit Function
                End If
            End If
        Next i
    Next Row
End Function

Private Function PropStr(shp As Visio.Shape, propName As String) As String
    On Error Resume Next
    If shp.CellExistsU("Prop." & propName, 0) Then
        PropStr = shp.CellsU("Prop." & propName).ResultStr("")
    End If
End Function

' True when a boolean Shape Data field (by internal name, else visible label) is
' set true. Tri-stateless: false or unset both read as False. Used for the
' indoor/outdoor flags (Prop.QInside / Prop.RInside).
Private Function PropIsTrue(shp As Visio.Shape, names As Variant) As Boolean
    On Error Resume Next
    Dim i As Long, Row As Long
    For i = LBound(names) To UBound(names)
        If shp.CellExistsU("Prop." & names(i), 0) Then PropIsTrue = TruthyCell(shp.CellsU("Prop." & names(i))): Exit Function
    Next i
    For Row = 0 To shp.RowCount(visSectionProp) - 1
        Dim lbl As String: lbl = shp.CellsSRC(visSectionProp, Row, 2).ResultStr(visNone)
        For i = LBound(names) To UBound(names)
            If lbl Like names(i) Then PropIsTrue = TruthyCell(shp.CellsSRC(visSectionProp, Row, 0)): Exit Function
        Next i
    Next Row
End Function
Private Function TruthyCell(c As Visio.Cell) As Boolean
    On Error Resume Next
    Dim s As String: s = UCase$(Trim$(c.ResultStr("")))
    If s = "TRUE" Or s = "1" Or s = "YES" Or s = "Y" Then TruthyCell = True: Exit Function
    If c.Result(visNone) <> 0 Then TruthyCell = True
End Function
' Tri-state boolean by internal name then visible label: 1 = true, 0 = false,
' -1 = the field isn't present at all (so the web app can apply its own default).
Private Function PropTri(shp As Visio.Shape, names As Variant) As Long
    On Error Resume Next
    PropTri = -1
    Dim i As Long, Row As Long
    For i = LBound(names) To UBound(names)
        If shp.CellExistsU("Prop." & names(i), 0) Then PropTri = IIf(TruthyCell(shp.CellsU("Prop." & names(i))), 1, 0): Exit Function
    Next i
    For Row = 0 To shp.RowCount(visSectionProp) - 1
        Dim lbl As String: lbl = shp.CellsSRC(visSectionProp, Row, 2).ResultStr(visNone)
        For i = LBound(names) To UBound(names)
            If lbl Like names(i) Then PropTri = IIf(TruthyCell(shp.CellsSRC(visSectionProp, Row, 0)), 1, 0): Exit Function
        Next i
    Next Row
End Function

' Fraction (0..1) of a queue exposed to sun/heat (Prop.QSun), by internal name
' then visible label. A bare "40" is read as 40%. Returns -1 when unset so the
' web app falls back to the qInside bool.
Private Function PropSunFrac(shp As Visio.Shape, names As Variant) As Double
    On Error Resume Next
    PropSunFrac = -1
    Dim i As Long, Row As Long
    For i = LBound(names) To UBound(names)
        If shp.CellExistsU("Prop." & names(i), 0) Then PropSunFrac = ClampFrac(shp.CellsU("Prop." & names(i))): Exit Function
    Next i
    For Row = 0 To shp.RowCount(visSectionProp) - 1
        Dim lbl As String: lbl = shp.CellsSRC(visSectionProp, Row, 2).ResultStr(visNone)
        For i = LBound(names) To UBound(names)
            If lbl Like names(i) Then PropSunFrac = ClampFrac(shp.CellsSRC(visSectionProp, Row, 0)): Exit Function
        Next i
    Next Row
End Function
Private Function ClampFrac(c As Visio.Cell) As Double
    On Error Resume Next
    ClampFrac = -1
    Dim s As String: s = Trim$(c.ResultStr(""))
    If s = "" Then Exit Function                 ' blank cell -> unset
    Dim v As Double: v = c.Result(visNone)
    If v > 1 Then v = v / 100                     ' allow "40" to mean 40%
    If v < 0 Then v = 0
    If v > 1 Then v = 1
    ClampFrac = v
End Function

Private Function CN(shp As Visio.Shape, cell As String) As Double
    On Error Resume Next
    CN = shp.CellsU(cell).ResultIU
End Function

'--------------------------- json builders ------------------------------------
Private Function NodeJson(id As String, isAttr As Boolean, x As Variant, y As Variant, nm As String) As String
    Dim s As String
    s = "  { ""id"": """ & id & """"
    If nm <> "" Then s = s & ", ""name"": """ & JStr(nm) & """"
    s = s & ", ""isAttraction"": " & LCase$(CStr(isAttr)) & _
        ", ""x"": " & CLng(x) & ", ""y"": " & CLng(y) & " }"
    NodeJson = s
End Function

Private Function AttractionJson(id As String, nm As String, entId As String, exId As String, _
                          x As Variant, y As Variant, ride As Double, cat As String, _
                          closed As Boolean, waitId As String, accessIds As String, _
                          hover As String, avgWait As Double, trackJson As String, queueJson As String, _
                          labelPosJson As String, thpw As String, audio As String, qInside As Boolean, rInsideTri As Long, _
                          qSun As Double) As String
    ' Emit each optional field only when set; otherwise lines match the original
    ' shape so the web app (ride/open defaults) is happy.
    Dim catJson As String
    If cat <> "" And cat <> "ride" Then catJson = ", ""category"": """ & cat & """"
    Dim closedJson As String
    If closed Then closedJson = ", ""closed"": true"
    Dim waitJson As String
    If waitId <> "" Then waitJson = ", ""waitId"": """ & JStr(waitId) & """"
    Dim thpwJson As String
    If Trim$(thpw) <> "" Then thpwJson = ", ""thpwId"": """ & JStr(Trim$(thpw)) & """"
    Dim accJson As String
    If accessIds <> "" Then accJson = ", ""accessNodeIds"": " & accessIds
    Dim hoverJson As String
    If Trim$(hover) <> "" Then hoverJson = ", ""hoverText"": """ & JStr(hover) & """"
    Dim avgJson As String
    If avgWait >= 0 Then avgJson = ", ""avgWait"": " & CLng(Round(avgWait))
    Dim trkJson As String
    If trackJson <> "" Then trkJson = ", ""track"": " & trackJson
    Dim queJson As String
    If queueJson <> "" Then queJson = ", ""queue"": " & queueJson
    Dim lblJson As String
    If labelPosJson <> "" Then lblJson = ", ""labelPos"": " & labelPosJson
    Dim audioJson As String
    If Trim$(audio) <> "" Then audioJson = ", ""audio"": """ & JStr(Trim$(audio)) & """"
    Dim insideJson As String       ' indoor flags for the sun/AC bar (only when set)
    If qSun >= 0 Then                                    ' fraction of the queue in the sun (hot head)
        Dim q As Double: q = Int(qSun * 1000 + 0.5) / 1000
        insideJson = insideJson & ", ""qSun"": " & Replace(CStr(q), ",", ".")
    ElseIf qInside Then insideJson = insideJson & ", ""qInside"": true"   ' legacy fully-shaded queue
    End If
    If rInsideTri = 1 Then insideJson = insideJson & ", ""rInside"": true"     ' indoor/AC (blue)
    If rInsideTri = 0 Then insideJson = insideJson & ", ""rInside"": false"    ' outdoor (yellow) — emitted so it beats the shop-indoor default
    AttractionJson = "  { ""id"": """ & id & """, ""name"": """ & JStr(nm) & _
        """, ""entranceNodeId"": """ & entId & """, ""exitNodeId"": """ & exId & _
        """, ""displayLocation"": { ""x"": " & CLng(x) & ", ""y"": " & CLng(y) & _
        " }, ""rideDuration"": " & CLng(Round(ride)) & catJson & closedJson & waitJson & thpwJson & accJson & hoverJson & avgJson & trkJson & queJson & lblJson & audioJson & insideJson & " }"
End Function

' ThemeParks.wiki entity GUID from Shape Data (Prop.ThPWID and aliases), used
' as the id to match standby waits + Lightning Lane. Tries internal name, then
' the visible label (Visio's NameU can differ from the label you typed).
' A real id is a GUID (non-numeric); an empty numeric cell reads back as
' "0.000", so a purely numeric value is treated as blank and skipped.
Private Function ThpwIdOf(shp As Visio.Shape) As String
    On Error Resume Next
    Dim names As Variant: names = Array("ThPWID", "ThpwId", "ThemeParksID", "TPWikiID")
    Dim i As Long, Row As Long, vstr As String
    For i = LBound(names) To UBound(names)
        If shp.CellExistsU("Prop." & names(i), 0) Then
            vstr = Trim$(shp.CellsU("Prop." & names(i)).ResultStr(""))
            If vstr <> "" And Not IsNumeric(vstr) Then ThpwIdOf = vstr: Exit Function
        End If
    Next i
    For Row = 0 To shp.RowCount(visSectionProp) - 1
        Dim lbl As String: lbl = shp.CellsSRC(visSectionProp, Row, 2).ResultStr(visNone)
        For i = LBound(names) To UBound(names)
            If lbl Like names(i) Then
                vstr = Trim$(shp.CellsSRC(visSectionProp, Row, 0).ResultStr(""))
                If vstr <> "" And Not IsNumeric(vstr) Then ThpwIdOf = vstr: Exit Function
            End If
        Next i
    Next Row
End Function

' Queue-Times ride id from the attraction's Shape Data (Prop.WaitID and aliases).
' Tries the internal name first, then falls back to matching the visible LABEL
' (Visio's internal NameU can differ from the label you typed). Numeric ids are
' normalized to a plain integer string so "284.00" matches the feed's "284".
Private Function WaitIdOf(shp As Visio.Shape) As String
    On Error Resume Next
    Dim names As Variant: names = Array("WaitID", "WaitId", "QueueID", "QueueTimesID", "QTID", "LiveID")
    Dim i As Long, Row As Long, vstr As String
    ' 1) by internal name: Prop.WaitID etc.
    For i = LBound(names) To UBound(names)
        If shp.CellExistsU("Prop." & names(i), 0) Then
            vstr = Trim$(shp.CellsU("Prop." & names(i)).ResultStr(""))
            If vstr <> "" Then WaitIdOf = NormalizeId(vstr): Exit Function
        End If
    Next i
    ' 2) by visible label (handles fields whose internal name was auto-generated)
    For Row = 0 To shp.RowCount(visSectionProp) - 1
        Dim lbl As String: lbl = shp.CellsSRC(visSectionProp, Row, 2).ResultStr(visNone)
        For i = LBound(names) To UBound(names)
            If lbl Like names(i) Then
                vstr = Trim$(shp.CellsSRC(visSectionProp, Row, 0).ResultStr(""))
                If vstr <> "" Then WaitIdOf = NormalizeId(vstr): Exit Function
            End If
        Next i
    Next Row
End Function

' Plain-integer string for numeric ids ("284.00" -> "284"); leave text as-is.
Private Function NormalizeId(ByVal s As String) As String
    If IsNumeric(s) Then NormalizeId = CStr(CLng(Val(s))) Else NormalizeId = s
End Function

' True when the shape's "Closed" shape data is set (not open at the park today).
' Accepts Visio Boolean cells (TRUE/FALSE) as well as text/numeric truthy values.
' Long-term closure flag. Look up Prop.Closed by internal name AND visible label,
' since a Shape Data field added in the UI usually lands in an auto-named row
' (Prop.Row_1) whose label — not its internal name — is "Closed".
Private Function IsClosed(shp As Visio.Shape) As Boolean
    IsClosed = PropIsTrue(shp, Array("Closed"))
End Function

Private Function JStr(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JStr = s
End Function

'--------------------------- small utilities ----------------------------------
Private Function NodeCenter(id As String) As Variant
    Dim v As Variant
    For Each v In mNodeMap
        If v(0) = id Then NodeCenter = Array(v(1), v(2)): Exit Function
    Next v
End Function

Private Function Dist2(a As Variant, b As Variant) As Double
    Dist2 = (a(0) - b(0)) ^ 2 + (a(1) - b(1)) ^ 2
End Function

Private Function ReverseCol(c As Collection) As Collection
    Dim r As New Collection, i As Long
    For i = c.Count To 1 Step -1: r.Add c(i): Next i
    Set ReverseCol = r
End Function

Private Function KeyExists(c As Collection, k As String) As Boolean
    ' IsObject() probes the item without Let-coercing it, so this works when the
    ' stored value is an object (e.g. a Collection) as well as a primitive.
    On Error GoTo nope
    Dim probe As Boolean: probe = IsObject(c(k))
    KeyExists = True
    Exit Function
nope:
    KeyExists = False
End Function

Private Sub Warn(msg As String)
    mWarnings.Add msg
End Sub

Private Function WriteOut(text As String) As String
    Dim folder As String
    On Error Resume Next
    folder = ThisDocument.path
    On Error GoTo 0
    If folder = "" Then folder = Environ$("USERPROFILE") & "\Desktop\"
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Dim path As String: path = folder & OUT_FILE
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
    WriteOut = path
End Function




