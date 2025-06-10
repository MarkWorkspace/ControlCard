-- Accessory script
-- see www.lua.org for Lua programming language
--
-- Extension functions - see manual
-- Globals:
-- gsScriptPath	Ends with \
-- gTblPlanks	Table of Frame plank guids having measure drawing or placed into active floor
-- gtblFrameMat	Table of material definitions from ArchiFrameMaterials.xml, key is mat id. Sorted by id.
-- gsManu		The manufacturer from xml-settings
-- Globals out:
-- gTblErr
-- If any operation fails, table is table of tables and cells are:
-- unid		Element's unique id (will be selected after listing)
-- id		ID of the element listed in result dump
-- pos      Position of error inside the plank in meters
-- text		The actual error message
local dummy, scriptDir

dummy, scriptDir = af_request("afpaths") -- From main data folder
dummy = nil

package.path = scriptDir .. "?.lua;" .. package.path
require("ArchiFrameCommon")

gScriptUtf8 = 1 -- We are fully utf

-- To set different info, set this to match with gTblPlanks. Fields: id. 
-- If saving element listing content is only here and different from other lists. In this case the table has structure:
-- key=index, value=tbl: index=for sorting, guid=plank guid, tblMaster=link to tblElemGuid2Master value for the master element of this multilayer, tblOwnerElem=direct owner having fields guid,type, elemgroup=usage in element (iElemGroup from plank object adjusted to something), elemgroupsort=sort key for the group, id=plank id
gtblPlanksData = nil

-- Globals
gXlsName = nil -- Name of the xls to be created

function GetFileExt(filename)
    return filename:match("^.+(%..+)$")
end
-- Unused after moving to Libxl
--[[
xlEdgeLeft = 7
xlEdgeRight = 10
xlEdgeBottom = 9
xlEdgeTop = 8
xlInsideHorizontal = 12
xlInsideVertical = 11

xlHairline = 1
xlMedium = -4138
xlThick = 4
xlThin = 2

xlPageBreakManual=-4135
xlHAlignLeft=-4131
xlHAlignRight=-4152
]]

-- ### COMMON DEFINITIONS/FUNCTIONS TO BE REMOVED IF USED IN CUSTOM CNC-WRITER BEG
-- # Frame machinings types
EMcFrAngledBegOld = 100
EMcFrAngledBeg = 101
EMcFrAngledBegTenon = 110 -- Also dovetail (OLD)
EMcFrBegHiddenShoe = 111
EMcFrAngledBegTenonMort = 112
EMcJointBeg = 113
EMcVCutBeg = 114
EMcFrAngledEndOld = 200
EMcFrAngledEnd = 201
EMcFrAngledEndTenon = 210
EMcFrEndHiddenShoe = 211
EMcFrAngledEndTenonMort = 212
EMcJointEnd = 213
EMcVCutEnd = 214

EMcFrOpening = 300 -- Opening - unused now
EMcFrGroove = 301 -- Logsin tapainen vapaa ura
EMcFrDrill = 302 -- Drilling
EMcFrMarking = 303 -- Marking
EMcFrReinforce = 304
EMcFrSaw = 305
EMcFrNailGroup = 306
EMcFrNailLine = 307

EMcFrTenonSide = 400 -- Also dovetail
EMcFrBalkJoint = 401 -- Narrowed balk

EMcLogTakasBalk = 900
EMcLogTakasLog = 901
EMcLogRounding = 902
EMcLogCutShape = 903
EMcFrTirol = 904

EMcFrBalkShoe = 1000
EMcFrHobaFix = 1001

EPS = 0.0001
MM2 = 0.0005
PI = 3.141592653589793
PI2 = PI * 0.5

gnCncErrCount = 0 -- Notices for the log

-- Adds cnc error message to gTblCncErr
function AddErrMsg(sGuid, sText)
    local sId

    if sText == nil then
        error("AddCncErr/sText")
    end

    if gnCncErrCount == 0 then
        gTblCncErr = {}
    end

    if sGuid ~= nil then
        sId = ac_getobjparam(sGuid, "#id")
    end

    gnCncErrCount = gnCncErrCount + 1
    gTblCncErr[gnCncErrCount] = {}
    gTblCncErr[gnCncErrCount].guid = sGuid
    gTblCncErr[gnCncErrCount].logid = sId
    gTblCncErr[gnCncErrCount].pos = 0
    gTblCncErr[gnCncErrCount].text = sText
end

-- COMMON FUNCTION IN A FEW PLACES IN DATA-FOLDER, UPDATE ALL IF UPDATING ONE
-- isHor  Is it horizontal structure (false/nil=default) (rotated inside the element, in walls true if horizontal framing)
-- lang   eng/swe/fin
-- elemgroup  Parameter iElemGroup
-- nFloorRoof, nil=not known or a wall, 1=it is a floor, 2=it is a roof
function GetElemGroupName(isHor, elemgroup, nFloorRoof)
    local lang

    -- ac_environment("tolog", string.format("%s", tostring(isFloor)))

    lang = af_request("aflang")
    if nFloorRoof and nFloorRoof == 1 then
        -- It is a floor
        if elemgroup == nil then
            elemgroup = ""
        elseif string.match(elemgroup, "^top.*") or string.match(elemgroup, "^2ndtop.*") or
            string.match(elemgroup, "^contour.*") or string.match(elemgroup, "^bottom.*") or
            string.match(elemgroup, "^2ndbottom.*") then
            if lang == "eng" then
                elemgroup = "Contour piece"
            elseif lang == "fin" then
                elemgroup = "Reunakappale"
            elseif lang == "swe" then
                elemgroup = "Contour piece"
            elseif lang == "nor" then
                elemgroup = "Kantbjelke"
            end
        else
            if lang == "eng" then
                elemgroup = "Joist"
            elseif lang == "fin" then
                elemgroup = "Vasa"
            elseif lang == "swe" then
                elemgroup = "Joist"
            elseif lang == "nor" then
                elemgroup = "Bjelke"
            end
        end
    elseif nFloorRoof and nFloorRoof == 2 then
        -- It is a roof
        if elemgroup == nil then
            elemgroup = ""
        elseif string.match(elemgroup, "^top.*") or string.match(elemgroup, "^2ndtop.*") or
            string.match(elemgroup, "^contour.*") or string.match(elemgroup, "^bottom.*") or
            string.match(elemgroup, "^2ndbottom.*") then
            if lang == "eng" then
                elemgroup = "Top/bottom"
            elseif lang == "fin" then
                elemgroup = "Ylä/alajuoksu"
            elseif lang == "swe" then
                elemgroup = "Top/bottom"
            elseif lang == "nor" then
                elemgroup = "Drager"
            end
        else
            if lang == "eng" then
                elemgroup = "Rafter"
            elseif lang == "fin" then
                elemgroup = "Vasa"
            elseif lang == "swe" then
                elemgroup = "Rafter"
            elseif lang == "nor" then
                elemgroup = "Sperr"
            end
        end
    elseif isHor then
        -- Horizontal structure
        if elemgroup == nil then
            elemgroup = ""
        elseif string.match(elemgroup, "^top%a*") or string.match(elemgroup, "^2ndtop%a*") or
            string.match(elemgroup, "^contour_x.*") then
            elemgroup = "Stud"
            if lang == "fin" then
                elemgroup = "Tolppa"
            elseif lang == "swe" then
                elemgroup = "Regel"
            elseif lang == "nor" then
                elemgroup = "Stendere"
            end
        elseif string.match(elemgroup, "^bottom%a*") or string.match(elemgroup, "^2ndbottom%a*") then
            elemgroup = "Stud"
            if lang == "fin" then
                elemgroup = "Tolppa"
            elseif lang == "swe" then
                elemgroup = "Regel"
            elseif lang == "nor" then
                elemgroup = "Stendere"
            end
        elseif string.match(elemgroup, "^vertical_x%a*") or string.match(elemgroup, "^vertical_y%a*") or
            string.match(elemgroup, "^contourtilted_opening%a*") then
            elemgroup = "Door/win"
            if lang == "fin" then
                elemgroup = "Ikk/ovi"
            elseif lang == "swe" then
                elemgroup = "Avväxling"
            elseif lang == "nor" then
                elemgroup = "Losholt" -- "Dør/vindu"
            end
        elseif string.match(elemgroup, "^vertical%a*") then
            elemgroup = "Hor"
            if lang == "fin" then
                elemgroup = "Vaaka"
            elseif lang == "nor" then
                elemgroup = "Vannrett"
            end
        elseif string.match(elemgroup, "^contourtilted%a*") then
            elemgroup = "Angled"
            if lang == "fin" then
                elemgroup = "Vinojuoksu"
            elseif lang == "swe" then
                elemgroup = "Regel"
            elseif lang == "nor" then
                elemgroup = "Vinklet"
            end
        elseif string.match(elemgroup, "^balk.*") then
            elemgroup = "Beam"
            if lang == "fin" then
                elemgroup = "Palkki"
            elseif lang == "swe" then
                elemgroup = "Bärlina"
            elseif lang == "nor" then
                elemgroup = "Bjelke"
            end
        elseif string.match(elemgroup, "^lintel.*") then
            elemgroup = "W/D beam"
            if lang == "fin" then
                elemgroup = "Aukkopalkki"
            elseif lang == "swe" then
                elemgroup = "Bärplanka"
            elseif lang == "nor" then
                elemgroup = "Dragere" -- "Dør-/vindu-bjelke"
            end
        elseif string.match(elemgroup, "^reinforce.*") then
            elemgroup = "Reinforcement"
            if lang == "fin" then
                elemgroup = "Vahvike"
            elseif lang == "swe" then
                elemgroup = "Förstärkning"
            elseif lang == "nor" then
                elemgroup = "Forsterkning"
            end
        elseif string.match(elemgroup, "^nogging.*") then
            elemgroup = "Nogging"
            if lang == "nor" then
                elemgroup = "Kubbing"
            elseif lang == "fin" then
                elemgroup = "Nurjahdustuki"
            elseif lang == "swe" then
                elemgroup = "Kortling"
            end
        elseif elemgroup == "sideleft_spacing" or elemgroup == "sideright_spacing" then
            elemgroup = "Side beams"
            if lang == "fin" then
                elemgroup = "Sivupalkki"
            elseif lang == "swe" then
                elemgroup = "Ändkortling"
            end
        else
            elemgroup = "" -- Just leave it empty, earlier was "? (elemgroup)"
        end
    else
        -- Vertical structure
        if elemgroup == nil then
            elemgroup = ""
        elseif string.match(elemgroup, "^top%a*") or string.match(elemgroup, "^2ndtop%a*") then
            elemgroup = "Top plate"
            if lang == "fin" then
                elemgroup = "Yläjuoksu"
            elseif lang == "swe" then
                elemgroup = "HB"
            elseif lang == "nor" then
                elemgroup = "Toppsvill"
            end
        elseif string.match(elemgroup, "^bottom%a*") or string.match(elemgroup, "^2ndbottom%a*") then
            elemgroup = "Bottom plate"
            if lang == "fin" then
                elemgroup = "Alajuoksu"
            elseif lang == "swe" then
                elemgroup = "Syll"
            elseif lang == "nor" then
                elemgroup = "Bunnsvill"
            end
        elseif string.match(elemgroup, "^contour_x.*") then
            elemgroup = "Top/bottom plate"
            if lang == "fin" then
                elemgroup = "Ylä/alajuoksu"
            elseif lang == "swe" then
                elemgroup = "HB/Syll"
            elseif lang == "nor" then
                elemgroup = "Topp-/Bunnsvill"
            end
        elseif string.match(elemgroup, "^vertical_x%a*") or string.match(elemgroup, "^contourtilted_opening%a*") then
            elemgroup = "Win/door"
            if lang == "fin" then
                elemgroup = "Ikk/ovi"
            elseif lang == "swe" then
                elemgroup = "Avväxling"
            elseif lang == "nor" then
                elemgroup = "Losholt" -- "Dør/vindu"
            end
        elseif string.match(elemgroup, "^vertical%a*") or string.match(elemgroup, "^contour_y.*") then
            elemgroup = "Stud"
            if lang == "fin" then
                elemgroup = "Tolppa"
            elseif lang == "swe" then
                elemgroup = "Regel"
            elseif lang == "nor" then
                elemgroup = "Stendere"
            end
        elseif string.match(elemgroup, "^contourtilted%a*") then
            elemgroup = "Angled"
            if lang == "fin" then
                elemgroup = "Vinojuoksu"
            elseif lang == "swe" then
                elemgroup = "Regel"
            elseif lang == "nor" then
                elemgroup = "Vinklet"
            end
        elseif string.match(elemgroup, "^balktop.*") then
            elemgroup = "Beam top"
            if lang == "fin" then
                elemgroup = "Palkki ylä"
            elseif lang == "swe" then
                elemgroup = "Bärlina topp"
            elseif lang == "nor" then
                elemgroup = "Bjelke"
            end
        elseif string.match(elemgroup, "^balkbot.*") then
            elemgroup = "Beam bottom"
            if lang == "fin" then
                elemgroup = "Palkki ala"
            elseif lang == "swe" then
                elemgroup = "Bärlina botten"
            elseif lang == "nor" then
                elemgroup = "Bjelke"
            end
        elseif string.match(elemgroup, "^lintel.*") then
            elemgroup = "W/D beam"
            if lang == "fin" then
                elemgroup = "Aukkopalkki"
            elseif lang == "swe" then
                elemgroup = "Bärplanka"
            elseif lang == "nor" then
                elemgroup = "Dragere" -- "Dør-/vindu-bjelke"
            end
        elseif string.match(elemgroup, "^nogging.*") then
            elemgroup = "Nogging"
            if lang == "nor" then
                elemgroup = "Kubbing"
            elseif lang == "fin" then
                elemgroup = "Nurjahdustuki"
            elseif lang == "swe" then
                elemgroup = "Kortling"
            end
        elseif elemgroup == "sideleft_spacing" or elemgroup == "sideright_spacing" then
            elemgroup = "Side beams"
            if lang == "fin" then
                elemgroup = "Sivupalkki"
            elseif lang == "swe" then
                elemgroup = "Ändkortling"
            end
        else
            elemgroup = "" -- Just leave it empty, earlier was "? (elemgroup)"
        end
    end

    return elemgroup
end

-- ### COMMON DEFINITIONS/FUNCTIONS TO BE REMOVED IF USED IN CUSTOM CNC-WRITER END

-- Gives file name to user specific folder or current script's folder
-- fnameNoExt	File name without language and file extension
-- Returns the full path file name
function XlsxGetTemplateFileName(fnameNoExt, fileExt)
	function GetFileExt(filename)
		return filename:match("^.+(%..+)$")
	end
    local apxPath, dataPath, userPath, fname, res

    apxPath, dataPath, userPath = af_request("afpaths")

    fname, res = XlsxGetTemplateFileNameInt(userPath, fnameNoExt, fileExt)
    if res then
        return fname
    end

    -- Fallback
    return XlsxGetTemplateFileNameInt(gsScriptPath, fnameNoExt, fileExt)
end

function XlsxGetTemplateFileNameInt(path, fnameNoExt, ext)
    local file, templateName, res

    res = false
    str = af.GetLangStr3()
    templateName = path .. fnameNoExt .. str .. ext
    file = io.open(templateName, "r")
    if file then
        io.close(file)
        res = true
    else
        templateName = path .. fnameNoExt .. "Eng" .. ext -- fallback

        file = io.open(templateName, "r")
        if file then
            io.close(file)
            res = true
        end
    end
    return templateName, res
end

---------------------------------------
-- DEFAULT LISTING

gnDefListType = 1 -- 1=ID per line, 2=group by plank size and length only, list related elements
gsElemIds = "" -- If prev is 2, this will have list of all element IDs

function DefaultOptions()
    local tblOptions, bRes, sErr, s

    tblOptions = {}

    s = "\"1:No grouping - single ID per line\",\"2:Group same length and list only element ID\""
    tblOptions[1] = {}
    tblOptions[1].cfgonly = 1
    tblOptions[1].type = 1
    tblOptions[1].prompt = "Grouping of planks"
    tblOptions[1].key = "group"
    tblOptions[1].defvalue = 1
    tblOptions[1].valuelist = s

    bRes, sErr = ac_optiondlg("FDDL", "Default listing options", tblOptions)
    if not bRes then
        if sErr ~= nil then
            af.RaiseError(sErr)
        end
        return false
    end

    gnDefListType = tblOptions[1].value
    return true
end

-- Called before save as dialog is showed
-- strPlnFileName	-> Full path file name of current pln WITHOUT extension (.pln removed)
-- Return values (multiple ret values):
-- 1=listing file name WITH extension, ""=Just prompt with file dlg, starting with *=result contains full path name - do not prompt
-- 2=file extension to be used in save as dialog
function OnInit(strPlnFileName)
    if not DefaultOptions() then
        gbCancel = true
        af.RaiseError("Canceled")
    end

    sFileName = strPlnFileName .. "_excel.xlsx"
    sExt = "xlsx"

    return sFileName, sExt
end

-- Called to do all that needs to be done
-- strFileName is full path name for the result file
function OnSaveListControlCard(strFileName, tblPlanksData)
    gXlsName = strFileName

    gtblPlanksData = tblPlanksData

    -- Added 3/2022: Group same length etc: Looks like gtblPlanksData was nil always
    if gnDefListType == 2 then
        local i, i2, v, t, s, tblElemIds

        -- Create own data table, sort it and save it back to gtblPlanks to get the order
        tblElemIds = {}
        gtblPlanksData = {}
        for i, v in ipairs(gTblPlanks) do
            t = {}

            ac_objectopen(v)
            t.guid = v

            t.info = af_request("plankinfo")
            t.id = "" -- Empty for external planks: ac_objectget("#id")
            if t.info.ownerelemid then
                t.id = t.info.ownerelemid
                if t.id == "" then
                    t.id = "(ELEM)"
                end

                i2 = 1
                while true do
                    if not tblElemIds[i2] then
                        tblElemIds[i2] = t.id
                        break
                    end
                    if tblElemIds[i2] == t.id then
                        break
                    end
                    i2 = i2 + 1
                end
            end

            t.matid = ac_objectget("iMatId")
            t.width, t.height = af.GetPlankSize()
            t.usage = "" -- Skip this one ac_objectget("iUsageId")
            t.len = af.GetPlankLength()

            -- Must use precision visible to the user - otherwise will have two different for lengths 1000,0 mm and 1000,1 mm
            t.sortwidth = ac_environment("ntos", t.width, "length", "dim")
            t.sortheight = ac_environment("ntos", t.height, "length", "dim")
            t.sortlen = ac_environment("ntos", t.len, "length", "dim")

            t.sortkey = string.format("%s %s %s", t.sortwidth, t.sortheight, t.sortlen) -- Length not reversed here - use only to see if needs new row in Excel. Not a real sortkey

            -- af.Log(string.format("id=%s t.sortwidth=%s t.sortheight=%s t.sortlen=%s key=%s", ac_objectget("#id"), t.sortwidth, t.sortheight, t.sortlen, t.sortkey))

            s = ac_objectget("iShowIDSep")
            ac_objectclose()

            -- We are collecting just planks in this listing - cut to the last ID separator
            if s then
            end

            gtblPlanksData[i] = t
        end

        -- Sort by: mat width, height, length longest first
        table.sort(gtblPlanksData, function(n1, n2)
            if n1.sortwidth ~= n2.sortwidth then
                return n1.width < n2.width
            end

            if n1.sortheight ~= n2.sortheight then
                return n1.height < n2.height
            end

            if n1.len ~= n2.len then
                return n1.len > n2.len -- Longest first
            end

            return false
        end)

        -- Sync gtblPlanks to be into order of gtblPlanksData
        gtblPlanks = {}
        for i, v in ipairs(gtblPlanksData) do
            gTblPlanks[i] = v.guid
        end

        -- Build list of element IDs
        table.sort(tblElemIds, function(s1, s2)
            return s1 < s2
        end)

        gsElemIds = ""
        for i, v in ipairs(tblElemIds) do
            if gsElemIds ~= "" then
                gsElemIds = gsElemIds .. ", "
            end
            gsElemIds = gsElemIds .. v
        end
    end

    local status, err = pcall(DoFrameExcel)

    af.LibxlClean()

    if not status then
        if excelStarted ~= nil then
            ws = nil
            wb = nil
            excelStarted:Quit()
            excelStarted = nil
        end
        af.RaiseError("Creating listing failed: " .. err)
    end
end

-- Called to parse tblPlank.elemgroup and tblPlank.elemgroupsort
-- strLang	nil=eng, fin supported
-- Returns plank's parameter group original val
function SetElemGroup(tblPlank, strLang)
    local group, groupOrg, sort, usage

    if strLang == nil then
        strLang = string.lower(af.GetLangStr3())
    end

    group = ac_objectget("iElemGroup")
    groupOrg = group

    group = GetElemGroupName(false, group)
    sort = nil

    if tblPlank.tblOwnerElem then
        if tblPlank.tblOwnerElem.type == "intstud" then
            sort = "1" .. group
            group = "Internal studding"
            if strLang == "fin" then
                group = "SP Koolaus"
            end
        end

        if tblPlank.tblOwnerElem.type == "extstud" or tblPlank.tblOwnerElem.type == "extstud2" then
            sort = "2" .. group
            group = "External studding"
            if strLang == "fin" then
                group = "UP Koolaus"
            end
        end
    end

    -- Override if iUsage not in default value ELEM
    usage = ac_objectget("iUsageId")
    if usage ~= "ELEM" and usage ~= "" then
        group = usage
    end

    if sort == nil then
        sort = "0" .. group
    end

    tblPlank.elemgroup = group
    tblPlank.elemgroupsort = sort -- To be able to sort differently from group name

    return groupOrg
end

---------------------------------------------------------------------
-- Special listings element list

function OnInitElem(strPlnFileName)
    sFileName = strPlnFileName .. "_elem.xlsx"
    sExt = "xlsx"

    return sFileName, sExt
end

function OnSaveListElem(strFileName)
    local i, v, tblElemGuid2Master, tblElemGuid2Owner, tblPlanksData, sId, plankinfo, elemMaster, elemOwner, tblPlank,
        nPlanks

    -- Do the grouping by master element here, for performance have a helper table: element guid -> table master element info (guid, id)

    tblElemGuid2Master = {} -- key=element guid, value=tbl with fields: guid=master element guid, id=its id for sorting, type=type attribute from xml <layer ref="WALL 42x42 VERT" ... type="intstud">
    tblElemGuid2Owner = {} -- As previous, but only guid, type
    tblPlanksData = {}

    nPlanks = 0
    for i, v in ipairs(gTblPlanks) do
        tblPlank = {}

        ac_objectopen(v)
        plankinfo = af_request("plankinfo")
        tblPlank.guid = v
        tblPlank.id = ac_objectget("#id")

        -- Skip if not part of an element
        if plankinfo.ownerelemguid then
            elemMaster = tblElemGuid2Master[plankinfo.ownerelemguid]
            elemOwner = tblElemGuid2Owner[plankinfo.ownerelemguid]

            if elemMaster == nil or elemOwner == nil then
                -- Open the parent element
                local elemParent, iElem, vElem, elemMasterTemp

                elemParent = af_request("elem_openparent", plankinfo.ownerelemguid)
                elemMaster = tblElemGuid2Master[elemParent.guid]

                elemMasterTemp = {}
                elemMasterTemp.guid = elemParent.guid

                elemOwner = {}
                elemOwner.guid = plankinfo.ownerelemguid

                for iElem, vElem in ipairs(elemParent.tblelems) do
                    if vElem.guid == elemParent.guid then
                        elemMasterTemp.type = vElem.type
                    end
                    if vElem.guid == plankinfo.ownerelemguid then
                        elemOwner.type = vElem.type
                    end
                end

                if elemMasterTemp.type == nil or elemOwner == nil then
                    af.RaiseError(string.format("Cannot find parent element for plank %s", v))
                end

                ac_objectclose()
                ac_objectopen(elemMasterTemp.guid)
                elemMasterTemp.id = ac_objectget("#id")
                elemMasterTemp.floor = ac_objectget("#floor")
                ac_objectclose()
                ac_objectopen(v)

                if elemMaster == nil then
                    elemMaster = elemMasterTemp
                end

                tblElemGuid2Master[elemParent.guid] = elemMaster
                tblElemGuid2Master[plankinfo.ownerelemguid] = elemMaster
                tblElemGuid2Owner[plankinfo.ownerelemguid] = elemOwner
            end

            nPlanks = nPlanks + 1
            tblPlank.index = nPlanks
            tblPlank.tblMaster = elemMaster
            tblPlank.tblOwnerElem = elemOwner
            SetElemGroup(tblPlank)

            tblPlanksData[nPlanks] = tblPlank
        end
        ac_objectclose()
    end

    -- Framen exceli
    gXlsName = strFileName
    gtblPlanksData = tblPlanksData
    local status, err = pcall(DoFrameExcelElem)

    if not status then
        if excelStarted ~= nil then
            ws = nil
            wb = nil
            excelStarted:Quit()
            excelStarted = nil
        end
        af.RaiseError("Creating listing failed: " .. err)
    end
end

-- Special listings element list
---------------------------------------------------------------------

---------------------------------------------------------------------
-- Special listings summary list

gsSummaryFinishLayer = ""
gtblSummaryLayers = nil -- Cache layer index->layer name

ESummaryInfoNone = 1 -- Column not used
ESummaryInfoColourName = 2 -- 3D material name
ESummaryInfoColourSample = 3 -- RGB to background
ESummaryInfoXlsQuality = 4 -- excel_quality
ESummaryInfoXlsType = 5 -- excel_type

gtblSummaryCols = nil -- What to put to columns, currently used from column I ([9] 1-based)

ESummaryCols = 12 -- Number of total columns, 1-based

function SummaryOptions()
    local tblOptions, bRes, sErr, s

    gtblSummaryLayers = {}
    tblOptions = {}
    tblOptions[1] = {}
    tblOptions[1].cfgonly = 1
    tblOptions[1].type = 4
    tblOptions[1].prompt = "Finishing layer in hotlinks (empty=from layer type)"
    tblOptions[1].key = "summary_finishlayer"
    tblOptions[1].defvalue = ""

    s =
        "\"1:Nothing\",\"2:3D surface material name\",\"3:3D surface material sample\",\"4:Quality (excel_quality from material list)\",\"5:Type (excel_type from material list)\""
    tblOptions[2] = {}
    tblOptions[2].cfgonly = 1
    tblOptions[2].type = 1
    tblOptions[2].prompt = "Column I content"
    tblOptions[2].key = "col_i"
    tblOptions[2].defvalue = 2
    tblOptions[2].valuelist = s

    tblOptions[3] = {}
    tblOptions[3].cfgonly = 1
    tblOptions[3].type = 1
    tblOptions[3].prompt = "Column J content"
    tblOptions[3].key = "col_j"
    tblOptions[3].defvalue = 3
    tblOptions[3].valuelist = s

    tblOptions[4] = {}
    tblOptions[4].cfgonly = 1
    tblOptions[4].type = 1
    tblOptions[4].prompt = "Column K content"
    tblOptions[4].key = "col_k"
    tblOptions[4].defvalue = 1
    tblOptions[4].valuelist = s

    tblOptions[5] = {}
    tblOptions[5].cfgonly = 1
    tblOptions[5].type = 1
    tblOptions[5].prompt = "Column L content"
    tblOptions[5].key = "col_l"
    tblOptions[5].defvalue = 1
    tblOptions[5].valuelist = s

    gHelpAnchor = "afdlglist_summary"
    bRes, sErr = ac_optiondlg("FDSO", "Summary listing options", tblOptions)
    gHelpAnchor = nil
    if not bRes then
        if sErr ~= nil then
            af.RaiseError(sErr)
        end
        return false
    end

    gsSummaryFinishLayer = string.lower(tblOptions[1].value)

    gtblSummaryCols = {}
    gtblSummaryCols[9] = tblOptions[2].value
    gtblSummaryCols[10] = tblOptions[3].value
    gtblSummaryCols[11] = tblOptions[4].value
    gtblSummaryCols[12] = tblOptions[5].value
    return true
end

-- Checks if currently opened object's layer is specified finishing layer
function SummaryIsFinish(guid)
    if gsSummaryFinishLayer == "" then
        return false
    end

    local t, attr

    t = ac_elemget(guid)
    if not t then
        return false
    end

    attr = gtblSummaryLayers[t.header.layer]
    if not attr then
        attr = ac_getattrinfo(2, t.header.layer)
        gtblSummaryLayers[t.header.layer] = attr
        attr.name = string.lower(attr.name)
        -- af.Log(string.format("attr.name=%s, gsSummaryFinishLayer=%s", attr.name, gsSummaryFinishLayer))
    end

    if string.match(attr.name, gsSummaryFinishLayer) then
        return true
    end
    return false
end

function OnInitSummary(strPlnFileName)
    local s, lang

    if not SummaryOptions() then
        gbCancel = true
        af.RaiseError("Canceled")
    end

    s = "summary"
    lang = af.GetLangStr3()
    if lang == "Nor" then
        s = "samleliste"
    elseif lang == "Fin" then
        s = "yhteenveto"
    end

    sFileName = strPlnFileName .. "_" .. s .. ".xlsx"
    sExt = "xlsx"

    return sFileName, sExt
end

function OnSaveListSummary(strFileName)
    local prevScript, prevCom

    prevScript = gScriptUtf8
    prevCom = ac_environment("luacomchars", 65001)

    gScriptUtf8 = 1
    local status, err = pcall(OnSaveListSummaryInt, strFileName)

    gScriptUtf8 = prevScript
    ac_environment("luacomchars", prevCom)
    gtblSummaryLayers = nil

    if not status then
        af.RaiseError(err)
    end

end

-- Also the plates, fills 
function GetSteelBeams(tblOthers)
    local tblSources, tblRes, k, v, tblElem, s, tblItem, tblNow

    tblRes = {}
    tblSources = ac_environment("getsel")
    if tblSources == nil then
        local status

        status = 2 + 4 -- APIFilt_OnVisLayer|APIFilt_OnActFloor
        if gnListingAll == 1 then
            status = 2 -- APIFilt_OnVisLayer
        end
        tblSources = ac_environment("getall", status, 6, 3) -- API_ObjectID, API_BeamID
        if tblSources == nil then
            ac_environment("tolog", "Old program version? Cannot get beams")
            return tblRes
        end
    end

    -- Collect any HEB beam or AC beam on layer "Stoldragere" or "Stalsoyler"
    for k, v in ipairs(tblSources) do
        tblElem = ac_elemget(v)

        -- Check layer ###

        tblItem = nil
        if tblElem.header.typeID == 6 then
            -- Check if HEB Beam
            ac_objectopen(v)

            s = string.lower(ac_objectget("#libname"))
            if string.find(s, "beam") or string.find(s, "bjelke") then
                tblItem = {}
                tblItem.matid = ac_objectget("Profile")
                if tblItem.matid == nil then
                    tblItem.matid = ac_objectget("stCurrentProfile")
                end
                tblItem.width = ac_objectget("B")
                tblItem.height = ac_objectget("zzyzx")
                tblItem.len = ac_objectget("A")
            elseif string.find(s, "archiframesteelplate") then
                tblItem = {}
                tblItem.matid = ac_objectget("iCode")
                tblItem.len = ac_objectget("A")
                tblItem.width = ac_objectget("B")
                tblItem.height = ac_objectget("zzyzx")
                tblItem.isplate = true
                if tblItem.matid == nil or tblItem.matid == "" then
                    tblItem.matid = string.format("%s x %s x %s", ac_environment("ntos", tblItem.len, "length", "dim"),
                        ac_environment("ntos", tblItem.width, "length", "dim"),
                        ac_environment("ntos", tblItem.height, "length", "dim"))
                end
            end

            ac_objectclose()
        end

        if tblItem and tblItem.matid then
            sKey = string.format("%s %06.3f x %07.3f x %07.3f", tblItem.matid, tblItem.width, tblItem.height,
                tblItem.len)

            tblNow = tblOthers[sKey]
            if tblNow == nil then
                tblNow = {}
                tblOthers[sKey] = tblNow
                tblNow.name = tblItem.matid
                tblNow.width = tblItem.width
                tblNow.height = tblItem.height
                tblNow.len = tblItem.len
                tblNow.num = 0
            end
            tblNow.num = tblNow.num + 1
        end
    end
end

-----------------------------------------------------------------------------
-- Edge Mcs shared in a few scripts

-- in:
--	mindex		Current index to iMcEdge
-- out, nil=edge not found, table having fields:
--	x1,y1,x2,y2	Edge line right pointing outside the board
--	dx,dy,len	Line vector and length
--	angle		Angle of the line 0...360
function FrMcEdgeFind(mindex)
    local contourbeg, row, n, edgeid

    -- Old param block? Ninth added 2021
    n = ac_objectget("iPolygon", 0, -1)
    if not n or n < 9 then
        return
    end

    contourbeg = 1
    row = 1
    edgeid = ac_objectget("iMcEdge", mindex, 2)
    if not edgeid then
        return
    end
    while true do
        n = ac_objectget("iPolygon", row, 9)
        if not n then
            break
        end
        if n == edgeid then
            local res

            -- Edge found, from this pt to next
            n = row + 1
            if ac_objectget("iPolygon", row, 3) == -1 then
                n = contourbeg
            end

            res = {}
            res.x1 = ac_objectget("iPolygon", row, 1)
            res.y1 = ac_objectget("iPolygon", row, 2)
            res.x2 = ac_objectget("iPolygon", n, 1)
            res.y2 = ac_objectget("iPolygon", n, 2)

            res.dx = res.x2 - res.x1
            res.dy = res.y2 - res.y1
            res.angle = math.atan2(res.dy, res.dx)
            res.len = math.sqrt(res.dx * res.dx + res.dy * res.dy)

            return res
        end
        if ac_objectget("iPolygon", row, 3) == -1 then
            contourbeg = row + 1 -- End of contour, end pt not duplicated (and no mask code from this pt to next)
        end

        row = row + 1
    end
end

-- Calcs screws etc.
function CalcMcQuantEdges(tblAcc)
    local i, rows, mc, tbl, s, edge

    rows = ac_objectget("iMcEdge", -1)
    if not rows then
        return
    end

    i = 1
    while i <= rows do
        mc = ac_objectget("iMcEdge", i, 1)
        if mc == 2000 then
            -- Groove
            -- double		dType;			// [1] EEdgeMcGroove
            -- double		dEdgeId;		// [2] edge id, iPolygon[row][9]
            -- double		dSide24;		// [3] Anchored either to front 2 or back 4 side
            -- double		dDepth;			// [4] From the edge towards inside (logical if at middle of the edge), if zero, will cut the side surface away always no matter to the angle
            -- double		dWidth;			// [5] If at surface this is the "depth" from side surface
            -- double		dOffFromSide;	// [6] Offset from side, negative=middle of the groove in the surface to the middle of the board edge
            -- double		dAngleCrossDeg;	// [7] Tilt angle difference from 90 degree ange: positive tilts groove inside the board, negative outside
            -- double		dExtBeg;		// [8] Extend groove at begin
            -- double		dExtEnd;		// [9] Extend groove at end

            s = ac_objectget("iMcEdgeStr", i)
            edge = FrMcEdgeFind(i)
            if s and s ~= "" and edge then
                -- Connection piece id given, calculate half of the edge length since comes to two boards

                len = edge.len + ac_objectget("iMcEdge", i, 8) + ac_objectget("iMcEdge", i, 9)
                if len > 0 then
                    tbl = tblAcc[s]
                    if tbl == nil then
                        tbl = {}
                        tbl.num = 0
                        tbl.len = 0
                        tblAcc[s] = tbl
                    end
                    tbl.len = tbl.len + len * 0.5 -- Half since connects to two
                    tbl.num = tbl.num + 0.5
                end

            end
        elseif mc == 2001 then
            -- Screws
            s = ac_objectget("iMcEdgeStr", i)
            edge = FrMcEdgeFind(i)
            if s and s ~= "" and edge then
                -- Screw type given, calculate number of screws

                -- Logic must be the same in GDL and here
                -- double		dDistFromEdge;	// 4 Distance from board edge
                -- double		dSpacing;		// 5 Spacing
                -- double		dDistFirst;		// 6 Distance of the first screw from edge begin (normal to the edge)
                -- double		dDistLast;		// 7 Distance of the lsat screw from edge end (normal to the edge)
                -- double		dAddLastTolerance;	// 8 Add last one as extra if further than this distance from spacing rule, negative=calculate even spacing not exceeding dSpacing so that first and last pos have screws, if negative, abs() of it is the min distance of the screws judging if short edge gets one or two screws

                local len, spacing, len, tolerance, count, d

                len = edge.len - ac_objectget("iMcEdge", i, 6) - ac_objectget("iMcEdge", i, 7)
                if len >= -0.001 then -- GDL uses same value
                    spacing = ac_objectget("iMcEdge", i, 5)
                    tolerance = ac_objectget("iMcEdge", i, 8)
                    if spacing < 0.001 then
                        spacing = 0.001
                    end

                    if tolerance < 0 then
                        -- Even spacing not exceeding given spacing
                        tolerance = -tolerance
                        if len < tolerance then
                            count = 1
                        else
                            count = math.floor(len / spacing)
                            d = len - count * spacing -- Unused space
                            if d > 0.001 then
                                count = count + 1
                            end
                            count = count + 1 -- Anyway there will be the last/first piece
                        end
                    else
                        -- Given spacing
                        count = math.floor(len / spacing)
                        d = len - count * spacing -- Unused space
                        if d > tolerance then
                            count = count + 1
                        end
                        count = count + 1 -- Anyway there will be the last/first piece
                    end

                    tbl = tblAcc[s]
                    if tbl == nil then
                        tbl = {}
                        tbl.num = 0
                        tblAcc[s] = tbl
                    end
                    tbl.num = tbl.num + count
                end

            end
        end
        i = i + 1
    end
end

-- Edge Mcs
-----------------------------------------------------------------------------

-- Calcs balk shoes etc
function CalcMcQuant(tblAcc)
    local i, count, mc, tbl, s, key, c

    count = ac_objectget("iMc", -1)
    for i = 1, count do
        mc = ac_objectget("iMc", i, 1)
        if mc == EMcFrBalkShoe then
            s = ac_objectget("iMcStr", i)
            if s and s ~= "" then
                tbl = tblAcc[s]
                if tbl == nil then
                    tbl = {}
                    tbl.num = 0
                    tblAcc[s] = tbl
                end

                local nBeg, nEnd, n

                nBeg = 1
                nEnd = 0
                n = ac_objectget("iMc", i, 2) -- flags
                if ac_environment("bittest", n, 0) == 1 then
                    nBeg = 0
                    nEnd = 1
                end
                if ac_environment("bittest", n, 3) == 1 then
                    nBeg = 1
                    nEnd = 1
                end

                tbl.num = tbl.num + nBeg + nEnd
            end
        elseif mc == EMcFrHobaFix then
            -- Hobafix
            s = ac_objectget("iMcStr", i)
            if s and s ~= "" then
                tbl = tblAcc[s]
                if tbl == nil then
                    tbl = {}
                    tbl.num = 0
                    tblAcc[s] = tbl
                end
                tbl.num = tbl.num + 1
            end
        elseif mc == EMcFrReinforce then
            -- Reinforcement found
            local height, len, width

            s = ac_objectget("iMcStr", i)
            height = ac_objectget("iMc", i, 5)
            len = ac_objectget("iMc", i, 4)
            width = ac_objectget("iMc", i, 6)
            if len == 0 then
                len = af.GetPlankLength()
            end

            c = 1
            if ac_objectget("iMc", i, 8) == 1 then
                -- Both sides
                c = 2
            end

            key = string.format("%s %06.3f x %07.3f x %07.3f", s, width, height, len)
            tbl = tblAcc[key]
            if tbl == nil then
                tbl = {}
                tbl.num = 0
                tbl.name = s
                tbl.width = width
                tbl.height = height
                tbl.len = len
                tblAcc[key] = tbl
            end

            tbl.num = tbl.num + c
        end
    end
    CalcMcQuantEdges(tblAcc)
end

function BoardIsInsu(guid)
    local code

    code = ac_getobjparam(guid, "iTypeCode")
    if code and code >= 200 and code < 300 then
        return true
    end
    return false
end

-- Returns true if any board in layer is flagged as insulation
function HasInsuBoards(tblboards)
    if tblboards == nil then
        return false
    end

    local i, code

    i = 1
    while tblboards[i] do
        if BoardIsInsu(tblboards[i].guid) then
            return true
        end
        i = i + 1
    end
    return false
end

-- Sets content of adjustable columns
function FillAdjCols(tblPlank, colname, colrgb)
    local k, v, adjcols, colval, matid, t

    adjcols = {}
    for k, v in pairs(gtblSummaryCols) do
        colval = nil
        if v == ESummaryInfoColourName then
            colval = colname
        elseif v == ESummaryInfoColourSample then
            colval = colrgb
        elseif (v == ESummaryInfoXlsQuality or v == ESummaryInfoXlsType) and tblPlank.plankinfo.type == 1 then
            matid = ac_objectget("iMatId")
            t = af_request("singlemat", matid)
            if t and t.xmlutf8 then
                local tag

                tag = "excel_quality"
                if v == ESummaryInfoXlsType then
                    tag = "excel_type"
                end

                colval = string.match(t.xmlutf8, string.format(" %s=\"(.-)\"", tag))
                -- af.Log(string.format("xml=%s\nATTR %s=%s", t.xmlutf8, tag, tostring(colval)))
            end
        end

        if colval then
            adjcols[k] = colval
            -- af.Log(string.format("KEY %s=%s", k, tostring(colval)))
        end
    end

    -- 4/2021: colname, colrgb not used any mor, adjcols is the new
    -- tblPlank.colname=colname
    -- tblPlank.colrgb=colrgb
    tblPlank.adjcols = adjcols
end

function OnSaveListSummaryInt(strFileName)
    local i, v, tblPlanksData, sId, plankinfo, tblPlank, nPlanks, tblHandledElems, tblElem, tblElemsWithPlanks,
        tblOthers, s, tbl

    -- Do the grouping by master element here, for performance have a helper table: element guid -> table master element info (guid, id)

    tblElemsWithPlanks = {} -- ArchiFrameElement-objects owning planks, value=true
    tblHandledElems = {} -- key=element guid, value=table having fields: elemtype
    tblPlanksData = {}
    tblOthers = {} -- key=part id, value=table having fields: num=quantity, type=len/vol (otherwise pieces), sortkey=optional - use instead of key if set, other optional fields:
    -- name, areanet, areagross, 
    -- To set corresponding fields in xls: width, height, len

    nPlanks = 0
    for i, v in ipairs(gTblPlanks) do
        tblPlank = {}

        ac_objectopen(v)
        plankinfo = af_request("plankinfo")
        tblPlank.guid = v
        tblPlank.id = ac_objectget("#id")

        CalcMcQuant(tblOthers)

        -- Calculate insulation from elements
        tblElem = nil
        if plankinfo.ownerelemguid then
            tblElem = tblHandledElems[plankinfo.ownerelemguid]
            if tblElem == nil then
                local q

                q = af_request("elem_quantities", plankinfo.ownerelemguid, 0)
                if q then
                    if q.quant.elemweightreduced > 0 and not HasInsuBoards(q.tblelems[1].tblboards) then
                        local insuName, dz
                        -- Calc only if it produces weight
                        insuName = "INSULATION"
                        if af.GetLangStr3() == "Fin" then
                            insuName = "ERISTE"
                        end
                        if q.geo.z1 then
                            dz = q.geo.z2 - q.geo.z1
                        else
                            -- No boards nor planks, take element's thickness'
                            local infoElem

                            infoElem = af_request("plankinfo", plankinfo.ownerelemguid)
                            dz = infoElem.width
                        end
                        s = string.format("%s %s", insuName, ac_environment("ntos", dz, "length", "dim"))
                        tbl = tblOthers[s]
                        if not tbl then
                            tbl = {}
                            tbl.sortkey = string.format("%s%6.2f", insuName, 1000 * (dz))
                            tbl.num = 0
                            tbl.areanet = 0
                            tbl.type = "vol"
                            tblOthers[s] = tbl
                        end

                        tbl.num = tbl.num + q.quant.elemvolumereduced
                        if dz > 0.001 then
                            tbl.areanet = tbl.areanet + q.quant.elemvolumereduced / dz
                        end
                    end

                    tblElem = {}
                    tblElem.elemtype = q.tblelems[1].elemtype
                    tblHandledElems[plankinfo.ownerelemguid] = tblElem
                end
            end
        end

        local handled, colname, colrgb

        handled = false
        if SummaryIsFinish(v) or (tblElem and string.match(tblElem.elemtype, "^finish.*")) then
            local dummy

            dummy, colname = ac_objectget("#material")
            colrgb = ac_objectget("#materialrgb")
        end

        if plankinfo.type == 4 then
            if ac_objectget("iPanelProfDX") > 0.001 then
                -- It is board having panels, explode and add each piece separately
                local panelblock

                panelblock = ac_objectget("iPanelBlock")
                ac_objectclose()

                local tblPtrs

                handled = true
                if panelblock ~= 4 then -- Number 4 is exploded to separate planks, just skip it
                    tblPtrs = af_request("mc_explodepanel", v, nil)
                    if tblPtrs then
                        local kpanel, vpanel, panelinfo

                        for kpanel, vpanel in ipairs(tblPtrs) do
                            ac_objectopen(vpanel)
                            panelinfo = af_request("plankinfo")

                            tblPlank = {}
                            tblPlank.guid = vpanel
                            tblPlank.id = ""
                            tblPlank.explodedowner = plankinfo.ownerelemguid

                            nPlanks = nPlanks + 1
                            tblPlank.index = nPlanks
                            tblPlank.plankinfo = panelinfo
                            tblPlank.len = af.GetPlankLength()
                            FillAdjCols(tblPlank, colname, colrgb)

                            tblPlanksData[nPlanks] = tblPlank
                            ac_objectclose()
                        end
                    end
                end
            end
        end

        if not handled then
            nPlanks = nPlanks + 1
            tblPlank.index = nPlanks
            tblPlank.plankinfo = plankinfo
            tblPlank.len = af.GetPlankLength()
            FillAdjCols(tblPlank, colname, colrgb)

            tblPlanksData[nPlanks] = tblPlank

            if plankinfo.ownerelemguid then
                tblElemsWithPlanks[plankinfo.ownerelemguid] = true
            end
            ac_objectclose()
        end
    end

    -- Remove any collected exploded planks if the element object owns any other plank
    local planks2, i2

    planks2 = {}
    i2 = 0
    for i = 1, nPlanks do
        v = tblPlanksData[i]
        if v.explodedowner == nil or tblElemsWithPlanks[v.explodedowner] == nil then
            -- Collect this one
            i2 = i2 + 1
            planks2[i2] = v
        end
    end
    tblPlanksData = planks2
    nPlanks = i2

    -- Add steel beams
    GetSteelBeams(tblOthers)

    -- Framen exceli
    gXlsName = ac_mbstoutf8(strFileName) -- Since file name was non-utf8
    gtblPlanksData = tblPlanksData
    status, err = pcall(DoFrameExcelSummary, tblOthers)

    if not status then
        af.RaiseError("Creating listing failed: " .. err)
    end
end

-- Special listings summary list
---------------------------------------------------------------------

---------------------------------------------------------------------
-- Special listings ArchiLogs window bucks

function OnInitBuck(strPlnFileName)
    sFileName = strPlnFileName .. "_bucks.xlsx"
    sExt = "xlsx"

    return sFileName, sExt
end

function OnFilterFuncBuck(sGuid)
    local sUsage

    ac_objectopen(sGuid)
    sUsage = ac_objectget("iUsageId")
    ac_objectclose()
    if sUsage == nil then
        return false
    end

    return string.upper(sUsage) == "BUCK"
end

-- Special listings ArchiLogs window bucks
---------------------------------------------------------------------

---------------------------------------------------------------------
-- Special listings cover planks

function OnInitCover(strPlnFileName)
    sFileName = strPlnFileName .. "_weather.xlsx"
    sExt = "xlsx"

    return sFileName, sExt
end

function OnFilterFuncCover(sGuid)
    local sUsage

    ac_objectopen(sGuid)
    sUsage = string.upper(ac_objectget("iUsageId"))
    ac_objectclose()
    if sUsage == nil then
        return false
    end

    if sUsage == "WEATHER-JAMB" then
        return false
    end

    return string.match(sUsage, "WEATHER%a*") ~= nil
end

function OnSaveListCover(strFileName)
    local i, v, sMatId, sUsage, sId, sOrgId, n, tblPlanksData, len, angle, nAngle, nHole, n2

    -- Override default IDs (forget about opening specific IDs since creating just a cut list)
    tblPlanksData = {}
    for i, v in ipairs(gTblPlanks) do
        tblPlanksData[i] = {}

        ac_objectopen(v)
        sMatId = ac_objectget("iMatId")

        sId = ac_objectget("#id")
        sOrgId = sId
        sUsage = ac_objectget("iUsageId")
        len = af.GetPlankLength()

        -- ac_msgbox(string.format("sId=%s sUsage=%s", sId, sUsage))

        -- Set id according to the angle
        angle = ac_objectget("iTiltAngle")
        if math.abs(angle) < PI / 180.0 then
            sId = "HOR"
        elseif math.abs(math.abs(angle) - PI2) < PI / 180.0 then
            sId = "VERT"
        else
            sId = "ANGLED"
        end

        -- Set usage according to the type
        if sUsage == "WEATHER-JAMB" then
            sUsage = "Jamb"
        elseif sUsage == "WEATHER-CORNER" then
            sUsage = "Corner"
            sId = ""
        elseif sUsage == "WEATHER-SUPP" then
            sUsage = "Support"
            sId = ""
        elseif sUsage == "WEATHER-COVER" then
            sUsage = "Covering"
            sId = ""
        elseif sUsage == "WEATHER-VISOR" then
            sUsage = "Visor"
        else
            sUsage = "Weather"
        end

        -- Add opening ID
        sId = sOrgId
        -- Already in the ID
        -- if string.match(sOrgId, "WB%-%a*")~=nil then
        --	sId=string.format("%s %s", sId, string.sub(sOrgId,4))
        -- end

        tblPlanksData[i].id = sId
        tblPlanksData[i].usage = sUsage

        --		if sMatId=="LA10010" then
        --			len=0.146
        --		elseif sMatId=="LA10015" then
        --			len=0.210
        --		else
        --			len=math.floor(len*100.0+1.49)/100.0		-- Minimum 5 mm extra
        --		end

        tblPlanksData[i].len = len
        ac_objectclose()
    end

    -- Save with default handler
    OnSaveListControlCard(strFileName, tblPlanksData)
end

-- Special listings cover planks
---------------------------------------------------------------------

---------------------------------------------------------------------
-- Special CSV elements

gCsvElemOnlyCore = false -- Added 3/2022: Include just core?

-- Called before collecting starts, will affect what iSpecialType(s) to collect
function OnPreCollectCsvElem()
    local tblOptions, bRes, sErr

    tblOptions = {}
    tblOptions[1] = {}
    tblOptions[1].cfgonly = 1
    tblOptions[1].type = 1
    tblOptions[1].prompt = "Handling of combined elements"
    tblOptions[1].key = "csv_spectype"
    tblOptions[1].defvalue = 1
    tblOptions[1].valuelist =
        "\"1:Include combined, skip sources (element split at building site)\",\"2:Include sources, skip combined (element split at factory)\""

    tblOptions[2] = {}
    tblOptions[2].cfgonly = 1
    tblOptions[2].type = 2
    tblOptions[2].prompt = "Include only core"
    tblOptions[2].key = "csv_onlycore"
    tblOptions[2].defvalue = 0

    bRes, sErr = ac_optiondlg("FDCV", "CSV saving options", tblOptions)
    if not bRes then
        if sErr ~= nil then
            af.RaiseError(sErr)
        end
        return -- Canceled
    end

    local t

    t = {}
    t.collectspecial = 2 -- Just combined
    if tblOptions[1].value == 2 then
        t.collectspecial = 1 -- Just sources
    end

    gCsvElemOnlyCore = false
    if tblOptions[2].value == 1 then
        gCsvElemOnlyCore = true
    end

    return t
end

function OnInitCsvElem(strPlnFileName)
    local s, lang

    lang = af.GetLangStr3()
    s = "_elementtransport.csv"
    if lang == "Fin" then
        s = "_elementtiluettelo.csv"
    end

    sFileName = strPlnFileName .. s
    sExt = "csv"

    return sFileName, sExt
end

function OnSaveListCsvElem(strFileName)
    local i, v, tblPlanksData, sId, plankinfo, elemMaster, nElems
    local tblElems -- 1-based table of all element parents, fields: id; dx, dy, dz for the size; weight; area; index
    local tblElemsHandled -- Processed element guids saved here
    local prevScript

    prevScript = gScriptUtf8
    gScriptUtf8 = 1

    tblElems = {}
    nElems = 0
    tblElemsHandled = {}
    for i, v in ipairs(gTblPlanks) do

        ac_objectopen(v)
        plankinfo = af_request("plankinfo")

        -- Skip if not part of an element
        if plankinfo.ownerelemguid and tblElemsHandled[plankinfo.ownerelemguid] == nil then
            -- Open the parent element
            local iElem, vElem, tblElem, q, guidCore

            tblElem = {}
            tblElem.id = "?"
            q = af_request("elem_quantities", plankinfo.ownerelemguid, 1)
            guidCore = nil
            for iElem, vElem in ipairs(q.tblelems) do
                tblElemsHandled[vElem.guid] = true
                if vElem.elemtype == "core" then
                    tblElem.id = vElem.id
                    guidCore = vElem.guid
                end
            end

            if gCsvElemOnlyCore and guidCore then
                q = af_request("elem_quantities", guidCore, 0)
            end

            tblElem.dx = q.geo.x2 - q.geo.x1
            tblElem.dy = q.geo.y2 - q.geo.y1
            tblElem.dz = q.geo.z2 - q.geo.z1
            tblElem.weight = q.quant.elemweight
            tblElem.area = tblElem.dx * tblElem.dy
            tblElem.index = nElems + 1
            tblElem.areanet = q.quant.areanet
            tblElem.areagross = q.quant.areagross

            nElems = nElems + 1
            tblElems[nElems] = tblElem
        end
        ac_objectclose()
    end

    -- CSV-muoto utf8-enkoodauksella, joka rivillä 5 saraketta
    local hFile, s

    hFile = io.open(strFileName, "wb")
    if not hFile then
        gScriptUtf8 = prevScript
        error("Opening file " .. strFileName .. " for writing failed. CHECK THAT FILE IS NOT OPEN IN EXCEL ETC.")
    end

    io.output(hFile)
    s = [[\xEF\xBB\xBF]]
    s = s:gsub("\\x(%x%x)", function(x)
        return string.char(tonumber(x, 16))
    end)
    io.write(s) -- utf8 bom-header

    -- Row 1
    io.write("PROJECT:;")
    io.write(af.Text2Csv(GetAutoTextNoNil("<PROJECTNAME>")))
    io.write(";;;;\r\n")

    io.write("NUMBER:;")
    io.write(af.Text2Csv(GetAutoTextNoNil("<PROJECTNUMBER>")))
    io.write(";;;;\r\n")

    io.write(";;;;;\r\n")
    io.write(";;;;;\r\n")

    s = "ELEMENT ID;LENGTH;HEIGHT;THICKNESS;WEIGHT;AREA M2;AREA GROSS;AREA NET\r\n"
    io.write(s)

    local status, err = pcall(SaveCsvElem, tblElems)
    io.close(hFile)

    gScriptUtf8 = prevScript

    if not status then
        af.RaiseError("Creating listing failed: " .. err)
    end
end

-- Compares:
-- 1. elem id
-- Returns: -1=n1<n2, 0=same, 1=n1>n2
function CmpElemKwCsv(n1, n2)
    -- Element id
    if n1.id < n2.id then
        return -1
    elseif n1.id > n2.id then
        return 1
    end

    return 0
end

function SaveCsvElem(tblElems)
    -- Sort by master element guid and plank index
    table.sort(tblElems, function(n1, n2)
        local cmp

        cmp = CmpElemKwCsv(n1, n2)
        if cmp ~= 0 then
            return cmp < 0
        end

        return n1.index < n2.index -- Keep the order if at same pos (id is part of sort order)
    end)

    local i, v, count, s

    i = 1
    while true do
        v = tblElems[i]
        if v == nil then
            break
        end
        io.write(af.Text2Csv(v.id))
        io.write(";")

        -- Len
        s = string.format("%.3f;", v.dx)
        s = ac_environment("strreplace", s, ".", ",")
        io.write(s)

        -- Height
        s = string.format("%.3f;", v.dy)
        s = ac_environment("strreplace", s, ".", ",")
        io.write(s)

        -- Height
        s = string.format("%.3f;", v.dz)
        s = ac_environment("strreplace", s, ".", ",")
        io.write(s)

        -- Weight
        s = string.format("%.0f;", v.weight)
        s = ac_environment("strreplace", s, ".", ",")
        io.write(s)

        -- Area
        s = string.format("%.2f;", v.area)
        s = ac_environment("strreplace", s, ".", ",")
        io.write(s)

        -- area grooss
        s = string.format("%.2f;", v.areagross)
        s = ac_environment("strreplace", s, ".", ",")
        io.write(s)

        -- area net
        s = string.format("%.2f;", v.areanet)
        s = ac_environment("strreplace", s, ".", ",")
        io.write(s)

        io.write("\r\n")

        i = i + 1
    end
end

-- Special CSV elements
---------------------------------------------------------------------

-- Special listings (default is added by ArchiFrame)
gtblListings = {}

gtblListings[1] = {}
gtblListings[1].strName = "Summary list"
gtblListings[1].strOnInitFunc = "OnInitSummary"
gtblListings[1].strOnSaveListFunc = "OnSaveListSummary"
gtblListings[1].nSorting = 2
gtblListings[1].nAllowSameId = 1
gtblListings[1].nCollectType = 5

gtblListings[2] = {}
gtblListings[2].strName = "Element listing"
gtblListings[2].strOnInitFunc = "OnInitElem"
gtblListings[2].strOnSaveListFunc = "OnSaveListElem" -- Use custom saver
gtblListings[2].nSorting = 2
gtblListings[2].nAllowSameId = 1
gtblListings[2].nCollectElem = 1

gtblListings[3] = {}
gtblListings[3].strName = "CSV Elements for transportation"
gtblListings[3].strPreCollectFunc = "OnPreCollectCsvElem"
gtblListings[3].strOnInitFunc = "OnInitCsvElem"
gtblListings[3].strOnSaveListFunc = "OnSaveListCsvElem" -- Use custom saver
gtblListings[3].nSorting = 2
gtblListings[3].nAllowSameId = 1
gtblListings[3].nCollectElem = 1
gtblListings[3].nCollectType = 5

gtblListings[4] = {}
gtblListings[4].strName = "ArchiLogs buck list"
gtblListings[4].strOnInitFunc = "OnInitBuck"
gtblListings[4].strOnSaveListFunc = "OnSaveList" -- Use default saver
gtblListings[4].strFilterFunc = "OnFilterFuncBuck"
gtblListings[4].nSorting = 2

gtblListings[5] = {}
gtblListings[5].strName = "Weatherboards"
gtblListings[5].strUIClass = "Cover"
gtblListings[5].strOnInitFunc = "OnInitCover"
gtblListings[5].strOnSaveListFunc = "OnSaveListCover" -- Use custom saver
gtblListings[5].strFilterFunc = "OnFilterFuncCover"
gtblListings[5].nSorting = 2
gtblListings[5].nAllowSameId = 1

gtblListings[6] = {}
gtblListings[6].strName = "Element amount of work listing"
gtblListings[6].strOnInitFunc = "OnInitElemWork"
gtblListings[6].strOnSaveListFunc = "OnSaveListElemWork"
gtblListings[6].nSorting = 2
gtblListings[6].nAllowSameId = 1
gtblListings[6].nCollectElem = 1

-----------------------------------------------------------------------------
-- Excelin kirjoitus

-- Creates excel from frame planks
function DoFrameExcel()
    local templateName, wb, ws
    local str, file, fileExt
    local projNumber, useNumToStr, res

    -- Do we have imperial working settings?
    af.CheckImperial()

    str = af.GetLangStr3()

    fileExt = af.GetFileExt(gXlsName)
    templateName = XlsxGetTemplateFileName("FrameListing", fileExt)
    if fileExt ~= ".xls" and fileExt ~= ".xlsx" then
        Planks2Txt()
        return
    end

    local book, sheet

    book = af.LibxlCreateBook(fileExt, templateName)
    sheet = book:get_sheet(0)

    Planks2Xls(book, sheet)

    if book:save(gXlsName) == false then
        af.RaiseError("Failed to save Excel workbook `" .. gXlsName ..
                          "`\nPlease make sure the file is not open in Excel already.")
    end

    book:release()

    -- Open the newly created one in excel
    af.ExcelOpen(gXlsName)
end

function GetColChar(colNum)
    return string.sub("ABCDEFGHIJKLMNOPQRSTUVWXYZ", colNum, colNum)
end

--[[
gDecSep=nil

function GetDecSep()
	if gDecSep==nil then
		gDecSep=gExcel.DecimalSeparator
	end
	return gDecSep
end

function NumToExcel( sNum )
	local	pos
	
	pos=string.find(sNum, ",", 1, true)
	if pos~=nil then
		sNum=string.sub(sNum,1,pos-1) .. '.' .. string.sub(sNum, pos+1,-1)
	end

	return sNum
end
]]

function GetAutoTextNoNil(autoName)
    local s

    s = ac_environment("parsetext", autoName)
    if s == nil then
        s = autoName
    end
    if s == "" then
        s = autoName
    end
    return s
end

-- Saves txt file if ws is null (mac or text output anyway)
function Planks2Xls(book, sheet)
    -- local colIds, colUsageId, colMatId, colWidth, colHeight, colCount, colLength, colSumLength, colTypeLen, colTypeM3, 
    local rowNum, rowNumTypeFirst, rowNumTblStart, bTypeChange
    local prevIdGroup, prevMatId, prevHeight, prevWidth, prevLen, prevMatLen, prevMatVol, prevUsage, cell
    local matLen, matInfo, val, bFlushInfo, bLastPlank, i, v, name, m3Factor
    local idNow, matId, height, width, angle, len, usage, similarCount, idList, totalCount, col
    local tblUsedIds = {}
    local i
    local idNowGroup -- Added 3/2022: May combine planks with diffent IDs into single line

    -- local tblCells = {}				-- Cells for single row (to be saved to xls/txt)
    -- local tblCellFormatTxt = {}		-- printf format strings if saving to txt file

    -- Header
    projName = GetAutoTextNoNil("<PROJECTNAME>")
    projNumber = GetAutoTextNoNil("<PROJECTNUMBER>")
    if projNumber == nil then
        cncName = "???.bvn"
    else
        cncName = projNumber .. ".bvn"
    end

    if book ~= nil and sheet ~= nil then
        af.LibxlMbsToCell(book, sheet, 0, 1, projName)
        af.LibxlMbsToCell(book, sheet, 1, 1, projNumber)
        af.LibxlMbsToCell(book, sheet, 2, 1, cncName)

        if gnDefListType == 2 and gsElemIds then
            af.LibxlMbsToCell(book, sheet, 0, 2, gsElemIds)
        end
    else
        io.write(string.format("PROJECT NAME:\t%s\n", projName))
        io.write(string.format("PROJECT NUMBER:\t%s\n", projNumber))
        io.write(string.format("CNC FILE:\t%s\n", cncName))
        io.write("Number\tName\tMat\tWidth\tHeight\tPieces\tLength\tLen tot\tType len\tType vol\n")
    end

    -- Where to write in txt-file only
    rowNum = 5 -- Starting row

    -- gTblPlanks is sorted by: matid, height, width, len, plankid
    prevIdGroup = "xxxNone**" -- Lisäys 8/2010: Samalla ID:llä samalle riville, jos ID vaihtuu, eri riville
    prevMatId = "xxxNone**"
    prevHeight = 0
    prevWidth = 0
    prevLen = 0
    prevMatLen = 0
    prevMatVol = 0
    prevUsage = ""
    idList = ""

    similarCount = 0
    similarLen = 0 -- Need to calc ourselves for imperial 
    totalCount = 0
    matLen = 0

    i = 1
    bLastPlank = false
    if gTblPlanks[1] == nil then
        return -- Special case - no planks
    end

    rowNumTypeFirst = rowNum
    rowNumTblStart = rowNum
    while true do
        -- Vikalla loopahduksella flushataan kerätty info eikä ole ole nyk kapulaa laisinkaan
        local idGroupIsSortkey -- true=do not compare lengths

        idGroupIsSortkey = nil
        v = gTblPlanks[i]
        if v == nil then
            bLastPlank = true
        else
            ac_objectopen(v)
            idNow = ac_objectget("#id")
            matId = ac_objectget("iMatId")
            width, height = af.GetPlankSize()
            usage = ac_objectget("iUsageId")
            angle = ac_objectget("iTiltAngle")
            len = af.GetPlankLength()
            ac_objectclose()

            idNowGroup = idNow
            if gtblPlanksData ~= nil and gtblPlanksData[i] ~= nil then
                -- Overrides?
                if gtblPlanksData[i].id ~= nil then
                    idNow = gtblPlanksData[i].id
                    idNowGroup = idNow
                end
                if gtblPlanksData[i].usage ~= nil then
                    usage = gtblPlanksData[i].usage
                end
                if gtblPlanksData[i].len ~= nil then
                    len = gtblPlanksData[i].len
                end
                if gtblPlanksData[i].matid ~= nil then
                    matId = gtblPlanksData[i].matid
                end
                if gtblPlanksData[i].width ~= nil then
                    width = gtblPlanksData[i].width
                end
                if gtblPlanksData[i].sortkey ~= nil then
                    idNowGroup = gtblPlanksData[i].sortkey
                    idGroupIsSortkey = true
                end
            end
        end

        bFlushInfo = false
        bTypeChange = false

        if matLen ~= 0 and
            (bLastPlank or
                (matId ~= prevMatId or math.abs(height - prevHeight) > af.EPS or math.abs(width - prevWidth) > af.EPS)) then
            -- Total for this material
            prevMatLen = matLen

            matInfo = gtblFrameMat[prevMatId]
            m3Factor = prevHeight * prevWidth
            if matInfo ~= nil then
                if matInfo.m3factor ~= nil and matInfo.m3factor ~= 0.0 then
                    m3Factor = matInfo.m3factor
                end
            end
            prevMatVol = prevMatLen * m3Factor

            if book ~= nil and sheet ~= nil then
                -- for col=1,colTypeM3 do
                -- sheet:set_format(rowNum - 1, col - 1).Borders(xlEdgeBottom).Weight = xlThin
                -- end
            end

            matLen = 0
            bFlushInfo = true
            bTypeChange = true
        end

        if (not idGroupIsSortkey and math.abs(len - prevLen) > 0.0005) or prevIdGroup ~= idNowGroup then
            bFlushInfo = true
        end

        if bFlushInfo and similarCount > 0 then
            -- Plank length changes, write all info now
            val = prevMatId
            name = prevMatId
            if val == "block" then
                -- Free size plank
                if af.CheckImperial() then
                    local widthStr, heightStr

                    widthStr = ac_environment("ntos", prevWidth, "length", "dim")
                    heightStr = ac_environment("ntos", prevHeight, "length", "dim")
                    val = string.format("%sx%s", widthStr, heightStr)
                else
                    val = string.format("%.0fx%.0f", prevWidth * 1000.0, prevHeight * 1000.0)
                end
            end

            if prevUsage == "" and gnDefListType ~= 2 then
                prevUsage = "general"
            end

            if book == nil or sheet == nil then
                -- To txt file
                local sLine

                sLine = string.format("%s\t%s\t%s\t", idList, prevUsage, val)
                sLine = sLine .. string.format("%s\t", ac_environment("ntos", prevWidth, "length", "dim"))
                sLine = sLine .. string.format("%s\t", ac_environment("ntos", prevHeight, "length", "dim"))
                sLine = sLine .. string.format("%s\t", string.format("%d", similarCount))
                sLine = sLine .. string.format("%s\t", ac_environment("ntos", prevLen, "length", "dim"))
                sLine = sLine .. string.format("%s\t", ac_environment("ntos", similarCount * prevLen, "length", "calc"))
                if bTypeChange then
                    sLine = sLine .. string.format("%s\t", ac_environment("ntos", prevMatLen, "length", "calc"))
                    sLine = sLine .. string.format("%s", ac_environment("ntos", prevMatVol, "volume", "calc"))
                end
                sLine = sLine .. "\n"

                io.write(sLine)

            else
                -- To xls file
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, idList, bTypeChange)
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, prevUsage, bTypeChange)
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 2, val, bTypeChange)
                af.LibxlDimLenToCell(book, sheet, rowNum - 1, 3, prevWidth, bTypeChange)
                af.LibxlDimLenToCell(book, sheet, rowNum - 1, 4, prevHeight, bTypeChange)
                af.LibxlNumToCell(book, sheet, rowNum - 1, 5, similarCount, bTypeChange)
                af.LibxlDimLenToCell(book, sheet, rowNum - 1, 6, prevLen, bTypeChange)
                af.LibxlCalcLenToCell(book, sheet, rowNum - 1, 7, similarCount * prevLen, bTypeChange)
                if bTypeChange then
                    af.LibxlCalcLenToCell(book, sheet, rowNum - 1, 8, prevMatLen, bTypeChange)
                    af.LibxlVolToCell(book, sheet, rowNum - 1, 9, prevMatVol, bTypeChange)
                end
            end

            totalCount = totalCount + similarCount
            similarCount = 0
            idList = ""
            tblUsedIds = {}
            rowNum = rowNum + 1
            if bTypeChange then
                rowNumTypeFirst = rowNum
            end
        end

        if bLastPlank then
            if book ~= nil and sheet ~= nil then
                -- cell=totalCount
                -- cell=string.format("=%s(%s%d:%s%d)", GetXlFuncName(EFuncSum), GetColChar(colCount), rowNumTblStart, GetColChar(colCount), rowNum-1)

                -- LibXL shouldn't need different names for Excel functions
                cell = string.format("SUM(F%d:F%d)", rowNumTblStart, rowNum - 1)
                sheet:write_formula(rowNum - 1, 5, cell) -- 5=column F

            else
                for i = 2, 6 do
                    io.write("\t")
                end

                io.write(string.format("%d\n", totalCount))
            end
            break
        end

        similarCount = similarCount + 1
        matLen = matLen + len

        prevMatId = matId
        prevHeight = height
        prevWidth = width
        prevLen = len
        prevIdGroup = idNowGroup
        prevUsage = usage
        if tblUsedIds[idNow] == nil then
            idList = idList .. idNow .. " "
            tblUsedIds[idNow] = 1
        end
        i = i + 1
    end
end

function Planks2Txt()
    local hFile

    hFile = io.open(gXlsName, "wt")
    if not hFile then
        error("Opening file " .. gXlsName .. " for writing failed. CHECK THAT FILE IS NOT OPEN IN EXCEL ETC.")
    end
    io.output(hFile)

    Planks2Xls(nil, nil)

    io.close(hFile)
end

-----------------------------------------------------------------------------
-- ELEMENT LIST

-- Compares:
-- 1. Master element id
-- 2. Grouping sort key
-- 3. Plank id
-- Returns: -1=n1<n2, 0=same, 1=n1>n2
function CmpElemPlank(n1, n2)
    -- Element id
    if n1.tblMaster.id < n2.tblMaster.id then
        return -1
    elseif n1.tblMaster.id > n2.tblMaster.id then
        return 1
    end

    -- Master element guid, nope, group just by visible id
    -- if n1.tblMaster.guid<n2.tblMaster.guid then
    --	return true
    -- elseif n1.tblMaster.guid>n2.tblMaster.guid then
    --	return false
    -- end

    -- Grouping text
    if n1.elemgroupsort < n2.elemgroupsort then
        return -1
    elseif n1.elemgroupsort > n2.elemgroupsort then
        return 1
    end

    return 0
end

function DoFrameExcelElem()
    local templateName
    local file, fileExt
    local projNumber, useNumToStr, res

    -- Sort by master element guid and plank index
    table.sort(gtblPlanksData, function(n1, n2)
        local cmp

        cmp = CmpElemPlank(n1, n2)
        if cmp ~= 0 then
            return cmp < 0
        end

        return n1.index < n2.index -- Keep the order if at same pos (id is part of sort order)
    end)

    -- Do we have imperial working settings?
    af.CheckImperial()

    fileExt = af.GetFileExt(gXlsName)
    templateName = XlsxGetTemplateFileName("FrameListingElem", fileExt)
    if fileExt ~= ".xls" and fileExt ~= ".xlsx" then
        Planks2TxtElem()
        return
    end

    local book = af.LibxlCreateBook(fileExt, templateName)
    local sheet = book:get_sheet(0)
    Planks2XlsElem(book, sheet)

    if not book:save(gXlsName) then
        error("Failed to save Excel workbook `" .. gXlsName ..
                  "`\nPlease make sure the file is not open in Excel already.")
    end

    book:release()

    af.ExcelOpen(gXlsName)

end

-- Saves txt file if ws is null (mac or text output anyway)
function Planks2XlsElem(book, sheet)
    local rowNum, addBreak, i1, i2, tblPlanks, nPlanks

    rowNum = 5 -- Starting row to the template
    addBreak = false
    i1 = 1
    while gtblPlanksData[i1] do
        -- Create table for planks in single elem
        i2 = i1
        nPlanks = 0
        tblPlanks = {}
        while true do
            if gtblPlanksData[i2] == nil then
                break
            end
            if gtblPlanksData[i1].tblMaster.id ~= gtblPlanksData[i2].tblMaster.id then
                break
            end

            nPlanks = nPlanks + 1
            tblPlanks[nPlanks] = gtblPlanksData[i2]
            i2 = i2 + 1
        end

        -- af.Log( string.format("i1=%d i2=%d", i1, i2) )
        -- dump(gtblPlanksData[i1].tblMaster)
        -- af.Log("")
        rowNum = Planks2XlsElemInt(gtblPlanksData[i1].tblMaster, book, sheet, rowNum, tblPlanks, addBreak)
        addBreak = true
        rowNum = rowNum + 1 -- One empty line

        i1 = i2
    end

end

function GetAngleStr(mcNum)
    local nRow, nCount, strRes, hasAngle, angleDeg

    hasAngle = false
    strRes = "90.0"
    nCount = ac_objectget("iMc", -1)
    for nRow = 1, nCount do
        if ac_objectget("iMc", nRow, 1) == mcNum then
            angleDeg = ac_objectget("iMc", nRow, 2)
            if math.abs(angleDeg - 90.0) < 0.1 then
                angleDeg = ac_objectget("iMc", nRow, 3)
            end
            if math.abs(angleDeg - 90.0) > 0.1 then
                if hasAngle then
                    strRes = "MANY"
                else
                    hasAngle = true
                    strRes = string.format("%.1f", angleDeg)
                end

            end
        end
    end

    return strRes
end

-- Saves txt file if book or sheet is null (text output anyway)
-- Writes planks for single element
-- rowNum		Number of first row to write
-- tblPlanks	Format as gtblPlanksData for element listing, planks for single element
-- addBreak		true=add page break before
-- Returns next unused row number
function Planks2XlsElemInt(tblMaster, book, sheet, rowNum, tblPlanks, addBreak)
    local prevElemgroup, colElemGroup
    local cell
    local matLen, matInfo, val, i, v, name
    local idNow, matId, height, width, len, usage, similarCount, idList, totalCount, col
    local i, scan
    local tblQuant, weight

    tblQuant = af_request("elem_quantities", tblMaster.guid, 1)
    -- dump(tblQuant)

    weight = tblQuant.quant.elemweight

    -- Header
    projName = string.format("%s, floor %d, element %s", GetAutoTextNoNil("<PROJECTNAME>"), tblMaster.floor,
        tblMaster.id)
    projNumber = GetAutoTextNoNil("<PROJECTNUMBER>")

    if book ~= nil and sheet ~= nil then
        if addBreak then
            -- ws.Cells(rowNum,1).PageBreak = xlPageBreakManual
            sheet:set_hor_page_break(rowNum - 1)
        end

        fmt = book:add_format()
        fmt:set_align_horizontal(1) -- ALIGNH_LEFT

        af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "Project:")
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, projName)
        sheet:set_format(rowNum - 1, 1, fmt)
        rowNum = rowNum + 1
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "Project number:")
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, projNumber)
        sheet:set_format(rowNum - 1, 1, fmt)
        rowNum = rowNum + 1
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "Client:")
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, GetAutoTextNoNil("<CLIENTNAME>"))
        sheet:set_format(rowNum - 1, 1, fmt)
        rowNum = rowNum + 1
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "Address 1:")
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, GetAutoTextNoNil("<CLIENTADDRESS1>"))
        sheet:set_format(rowNum - 1, 1, fmt)
        rowNum = rowNum + 1
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "Address 2:")
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, GetAutoTextNoNil("<CLIENTADDRESS2>"))
        sheet:set_format(rowNum - 1, 1, fmt)
        rowNum = rowNum + 1
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "Engineer:")
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, GetAutoTextNoNil("<CADTECHNICIAN>"))
        sheet:set_format(rowNum - 1, 1, fmt)
        rowNum = rowNum + 1
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "Element weight:")
        af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, string.format("%.0f kg", weight)) -- Element weight
        sheet:set_format(rowNum - 1, 1, fmt)
        rowNum = rowNum + 1
    else
        if addBreak then
            io.write("\f")
        end

        io.write(string.format("PROJECT NAME:\t%s\n", projName))
        io.write(string.format("PROJECT NUMBER:\t%s\n", projNumber))
        io.write(string.format("CLIENT\t%s\n", GetAutoTextNoNil("<CLIENTNAME>")))
        io.write(string.format("ADDRESS1:\t%s\n", GetAutoTextNoNil("<CLIENTADDRESS1>")))
        io.write(string.format("ADDRESS1:\t%s\n", GetAutoTextNoNil("<CLIENTADDRESS1>")))
        io.write(string.format("ENGINEER:\t%s\n", GetAutoTextNoNil("<CADTECHNICIAN>")))
        io.write(string.format("ELEMENT WEIGHT:\t%.0f kg\n", weight))
    end

    i = 1
    bLastPlank = false
    if tblPlanks[1] == nil then
        return rowNum -- Special case - no planks
    end

    prevElemgroup = "xxx***#%%&" -- Write header always when changed

    while true do
        v = tblPlanks[i]
        if v == nil then
            break
        end

        -- Number of similar planks
        scan = i + 1
        while tblPlanks[scan] ~= nil do
            if tblPlanks[i].id ~= tblPlanks[scan].id or CmpElemPlank(tblPlanks[i], tblPlanks[scan]) ~= 0 then
                break
            end
            scan = scan + 1
        end

        if tblPlanks[i].id == "" then
            af.Log(string.format("WARNING: Element %s has plank(s) without ID, please update the element", tblMaster.id))
        end

        ac_objectopen(v.guid)
        idNow = ac_objectget("#id")
        matId = ac_objectget("iMatId")
        width, height = af.GetPlankSize()
        len = af.GetPlankLength()

        if matId == "block" or matId == "plane" or matId == "round" then
            -- Free size plank
            if af.CheckImperial() then
                local widthStr, heightStr

                widthStr = ac_environment("ntos", width, "length", "work")
                heightStr = ac_environment("ntos", height, "length", "work")
                matId = string.format("%sx%s", widthStr, heightStr)
            else
                matId = string.format("%.0fx%.0f", width * 1000.0, height * 1000.0)
            end
        end

        local angleBeg, angleEnd

        angleBeg = GetAngleStr(101)
        angleEnd = GetAngleStr(201)

        --[[
		tblCells[colIds]		= idNow
		tblCells[colMatId]		= matId
		tblCells[colLength]		= strLen

		tblCells[colAngleBeg]	= angleBeg
		tblCells[colAngleEnd]	= angleEnd

		tblCells[colCount]		= scan-i
		tblCells[colElemGroup]	= v.elemgroup
		]]

        ac_objectclose()

        if book == nil or sheet == nil then
            -- To txt file
            local sLine
            if prevElemgroup ~= v.elemgroup then
                prevElemgroup = v.elemgroup
                io.write(string.format("\n%s\nPlank id\tMaterial\tLength\tAngle beg\tAngle end\tCount\tType\n",
                    v.elemgroup))
            end

            sLine = string.format("%s\t%s\t%s\t", idNow, matId, ac_environment("ntos", len, "length", "dim"))
            sLine = sLine .. string.format("%s\t%s\t%d\t%s\n", angleBeg, angleEnd, scan - i, v.elemgroup)

            io.write(sLine)

        else
            if prevElemgroup ~= v.elemgroup then
                prevElemgroup = v.elemgroup
                rowNum = rowNum + 1

                af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, v.elemgroup)

                rowNum = rowNum + 1

                af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "Plank id")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, "Material")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 2, "Length")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 3, "Angle beg")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 4, "Angle end")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 5, "Count")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 6, "Type")

                rowNum = rowNum + 1
            end

            -- To xls file
            af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, idNow)
            af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, matId)
            af.LibxlDimLenToCell(book, sheet, rowNum - 1, 2, len)
            af.LibxlMbsToCell(book, sheet, rowNum - 1, 3, angleBeg)
            af.LibxlMbsToCell(book, sheet, rowNum - 1, 4, angleEnd)
            af.LibxlNumToCell(book, sheet, rowNum - 1, 5, scan - i)
            af.LibxlMbsToCell(book, sheet, rowNum - 1, 6, v.elemgroup)
        end

        rowNum = rowNum + 1
        i = scan
    end

    -- Boards to the end
    local nElem, vElem
    local widthStr, heightStr

    for nElem, vElem in ipairs(tblQuant.tblelems) do
        if vElem.tblboards then
            -- non-nil
            if book == nil or sheet == nil then
                -- To txt file
                io.write(string.format("\n%s\nBoard id\tWidth\tHeight\tCount\tNet area\n", vElem.elemtypeid))
            else
                rowNum = rowNum + 1
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, vElem.elemtypeid)

                rowNum = rowNum + 1

                -- bold!
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, "ID")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 1, "Width")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 2, "Height")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 3, "Count")
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 4, "Net area")

                rowNum = rowNum + 1
            end

            table.sort(vElem.tblboards, function(n1, n2)
                return n1.id < n2.id
            end)

            i = 1
            while vElem.tblboards[i] do
                scan = i + 1
                while vElem.tblboards[scan] do
                    if vElem.tblboards[i].id ~= vElem.tblboards[scan].id then
                        break
                    end

                    if math.abs(vElem.tblboards[i].width - vElem.tblboards[scan].width) > 0.001 or
                        math.abs(vElem.tblboards[i].height - vElem.tblboards[scan].height) > 0.001 then
                        af.Log(string.format(
                            "ERROR: boards with same ID (%s) different size, update element and create listing again",
                            vElem.tblboards[i].id))
                        break
                    end

                    scan = scan + 1
                end

                -- dump(vElem.tblboards[i])
                local width, height

                width = vElem.tblboards[i].width
                height = vElem.tblboards[i].height

                -- Check if panel instead of board
                local idStr, area

                idStr = vElem.tblboards[i].id
                ac_objectopen(vElem.tblboards[i].guid)
                if ac_objectget("iPanelProfDX") > 0.0005 then
                    idStr = ac_objectget("iMatId") -- Instead of element type ID
                end
                area = ac_objectget("iArea")
                ac_objectclose()

                if book == nil or sheet == nil then
                    -- To txt file
                    io.write(string.format("%s\t%s\t%s\t%d\t%.3f\n", idStr,
                        ac_environment("ntos", width, "length", "dim"), ac_environment("ntos", height, "length", "dim"),
                        scan - i, (scan - i) * area))
                else
                    af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, idStr)

                    af.LibxlDimLenToCell(book, sheet, rowNum - 1, 1, width)
                    af.LibxlDimLenToCell(book, sheet, rowNum - 1, 2, height)
                    af.LibxlNumToCell(book, sheet, rowNum - 1, 3, scan - i)
                    af.LibxlAreaToCell(book, sheet, rowNum - 1, 4, (scan - i) * area)

                    rowNum = rowNum + 1
                end

                i = scan
            end
        end
    end

    return rowNum
end

function Planks2TxtElem()
    local hFile

    hFile = io.open(gXlsName, "wt")
    if not hFile then
        error("Opening file " .. gXlsName .. " for writing failed. CHECK THAT FILE IS NOT OPEN IN EXCEL ETC.")
    end
    io.output(hFile)

    Planks2XlsElem(nil)

    io.close(hFile)
end

-----------------------------------------------------------------------------
-- SUMMARY LIST

function CmpMatId(n1, n2)
    if n1.plankinfo.typename < n2.plankinfo.typename then
        return -1
    end
    if n1.plankinfo.typename > n2.plankinfo.typename then
        return 1
    end
    return 0
end

function CmpAdjCols(n1, n2)
    local s1, s2, col

    -- Col I/9 is the first adjustable
    for col = 9, ESummaryCols do
        s1 = n1.adjcols[col]
        if s1 == nil then
            s1 = ""
        else
            s1 = tostring(s1)
        end
        s2 = n2.adjcols[col]
        if s2 == nil then
            s2 = ""
        else
            s2 = tostring(s2)
        end
        if s1 < s2 then
            return -1
        end
        if s1 > s2 then
            return 1
        end
    end

    return 0

    --[[
	s1=n1.colname
	if s1==nil then
		s1=""
	end
	s2=n2.colname
	if s2==nil then
		s2=""
	end
	if s1<s2 then
		return -1
	end
	if s1>s2 then
		return 1
	end
	return 0
]]
end

-- Returns: -1=n1<n2, 0=same, 1=n1>n2
function CmpSummaryPlank(n1, n2)
    local d

    -- Planks first
    d = n1.plankinfo.type - n2.plankinfo.type
    if d ~= 0 then
        return d
    end

    if n1.plankinfo.type == 4 then
        -- Boards just by matid
        return CmpMatId(n1, n2)
    end

    -- Planks by plank size
    d = n1.plankinfo.width - n2.plankinfo.width
    if math.abs(d) > af.EPS then
        return d
    end

    d = n1.plankinfo.height - n2.plankinfo.height
    if math.abs(d) > af.EPS then
        return d
    end

    d = CmpMatId(n1, n2)
    if d ~= 0 then
        return d
    end

    return CmpAdjCols(n1, n2)
end

function DoFrameExcelSummary(tblOthers)
    local templateName, wb, ws
    local str, fileExt
    local projNumber, useNumToStr, res

    -- Sort by master element guid and plank index
    table.sort(gtblPlanksData, function(n1, n2)
        local cmp

        cmp = CmpSummaryPlank(n1, n2)
        if cmp ~= 0 then
            return cmp < 0
        end

        return n1.index < n2.index -- Keep the order if at same pos (id is part of sort order)
    end)

    -- Do we have imperial working settings?
    af.CheckImperial()

    fileExt = af.GetFileExt(gXlsName)
    templateName = XlsxGetTemplateFileName("FrameListingSummary", fileExt)
    if fileExt ~= ".xls" and fileExt ~= ".xlsx" then
        Planks2TxtSummary(tblOthers)
        return
    end

    --
    -- NOTE libxl interface doesn't convert indices to Lua convention (1...N), but uses regular (0...N-1)
    -- 

    local book = af.LibxlCreateBook(fileExt, templateName)
    local sheet = book:get_sheet(0)

    Planks2XlsSummary(book, sheet, tblOthers)

    if not book:save(gXlsName) then
        error("Failed to save Excel workbook `" .. gXlsName ..
                  "`\nPlease make sure the file is not open in Excel already.")
    end

    book:release()

    af.ExcelOpen(gXlsName)
end

-- Saves txt file if ws is null (mac or text output anyway)
function Planks2XlsSummary(book, sheet, tblOthers)
    local rowNum, i1, i2, tblPlanks, nPlanks, projName, projNumber

    projName = GetAutoTextNoNil("<PROJECTNAME>")
    projNumber = GetAutoTextNoNil("<PROJECTNUMBER>")

    if sheet ~= nil then
        af.LibxlMbsToCell(book, sheet, 0, 1, projName)
        af.LibxlMbsToCell(book, sheet, 1, 1, projNumber)

        -- Added 4/2021: adjustable columns - set headers
        local wsHelp, s

        wsHelp = book:get_sheet(1) -- Second shet
        if wsHelp then
            -- Col I/9 is the first adjustable
            for i1 = 9, ESummaryCols do
                i2 = gtblSummaryCols[i1]
                s = ""
                if i2 ~= ESummaryInfoNone then
                    s = wsHelp:read_str(37 + i2, 0) -- in Excel row 40 is the first header row
                    -- af.Log(string.format("wsHelp: i2=%s, s=%s", tostring(i2), tostring(s)))
                    af.LibxlMbsToCell(book, sheet, 2, i1 - 1, tostring(s))
                end
            end
        end

    else
        io.write(string.format("PROJECT NAME:\t%s\n", projName))
        io.write(string.format("PROJECT NUMBER:\t%s\n", projNumber))
        io.write(string.format(
            "Material ID\tWidth\tHeight\tPcs\tTotal len\tArea gross\tArea net\tTotal vol gr\tColour\tColour2\n",
            projNumber))
    end

    rowNum = 4 -- Starting row to the template
    i1 = 1

    if gtblPlanksData[i1] == nil then
        -- Just other stuff, no planks

        Planks2XlsSummaryInt(book, sheet, rowNum, gtblPlanksData, tblOthers)
        return
    end

    while gtblPlanksData[i1] do
        -- Find next different piece
        local othersNow

        othersNow = nil
        i2 = i1
        nPlanks = 0
        tblPlanks = {}
        while true do
            if gtblPlanksData[i2] == nil then
                othersNow = tblOthers
                break
            end
            if CmpSummaryPlank(gtblPlanksData[i1], gtblPlanksData[i2]) ~= 0 then
                if i1 == i2 then
                    i2 = i2 + 1
                    af.Log("Bug in listing") -- Should never occur
                end
                break
            end

            nPlanks = nPlanks + 1
            tblPlanks[nPlanks] = gtblPlanksData[i2]
            i2 = i2 + 1
        end
        rowNum = Planks2XlsSummaryInt(book, sheet, rowNum, tblPlanks, othersNow)
        i1 = i2
    end

end

-- Saves txt file if ws is null (mac or text output anyway)
-- Writes planks for single element
-- rowNum		Number of first row to write
-- tblPlanks	Format as gtblPlanksData for element listing, planks for single element
-- othersNow	Other collected items (table key=id, value fields: num=quantity)
-- Returns next unused row number
function Planks2XlsSummaryInt(book, sheet, rowNum, tblPlanks, othersNow)
    local cell
    local matLen, matInfo, val, i, i1, v, name, strLen
    local col
    local scan

    --[[
	colMatId	=1
	colWidth	=2
	colHeight	=3
	colCount	=4
	colTotLen	=5
	colTotAreaGross	=6
	colTotAreaNet	=7
	colTotVol	=8
	colColour	=9
	colColourRgb=10
	]]

    i1 = 1
    while true do
        if tblPlanks[i1] == nil then
            break
        end

        -- Number of planks from the same mat
        local matId, width, height, totLen, totAreaGross, totAreaNet, totVol, currplank

        v = tblPlanks[i1]
        currplank = v
        matId = v.plankinfo.typename
        width = 0
        height = 0
        totLen = 0
        totAreaGross = 0
        totAreaNet = 0
        totVol = 0

        scan = i1
        while true do
            v = tblPlanks[scan]
            if v == nil or CmpSummaryPlank(tblPlanks[i1], v) ~= 0 then
                break
            end

            if v.plankinfo.type == 4 then
                -- Board
                local poly, dx, dy

                poly = af_request("getpoly", nil, v.plankinfo.guid)
                dx = poly.x2 - poly.x1
                dy = poly.y2 - poly.y1

                if BoardIsInsu(v.plankinfo.guid) then
                    -- Insulation object, get exact net volume
                    local info2, tblSettings

                    tblSettings = {}
                    tblSettings.quant = 2
                    info2 = af_request("plankinfo", v.plankinfo.guid, nil, tblSettings)
                    if info2.width > width then
                        width = info2.width
                    end
                    totAreaGross = totAreaGross + poly.area -- For insulation the gross area is polygon of insulation object not reducing framing
                    totAreaNet = totAreaNet + info2.acvol / info2.width -- info2.width=insulation thickness
                    totVol = totVol + poly.area * info2.width -- Calc as gross vol
                else
                    -- Board or cladding
                    if dx > width then
                        width = dx
                    end
                    if dy > height then
                        height = dy
                    end
                    totAreaGross = totAreaGross + dx * dy
                    totAreaNet = totAreaNet + poly.area
                    totVol = totVol + dx * dy * v.plankinfo.width
                end
            else
                -- Plank
                width = v.plankinfo.width
                height = v.plankinfo.height
                totLen = totLen + v.len
                totVol = totVol + width * height * v.len

                totAreaGross = totAreaGross + v.len * v.plankinfo.height
            end
            scan = scan + 1
        end

        if sheet == nil then
            -- To txt file
            local sLine

            sLine = string.format("%s\t%s\t%s\t", matId, ac_environment("ntos", width, "length", "dim"),
                ac_environment("ntos", height, "length", "dim"))
            sLine = sLine .. string.format("%d\t", scan - i1)
            sLine = sLine .. string.format("%s\t", ac_environment("ntos", totLen, "length", "calc"))
            sLine = sLine .. string.format("%s\t", ac_environment("ntos", totAreaGross, "area", "calc"))
            sLine = sLine .. string.format("%s\t", ac_environment("ntos", totAreaNet, "area", "calc"))
            sLine = sLine .. string.format("%s\t", ac_environment("ntos", totVol, "volume", "calc"))
            sLine = sLine .. string.format("%s\t\n", colour)

            -- TODO
            io.write(sLine)

        else
            -- To xls file
            af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, matId)
            af.LibxlDimLenToCell(book, sheet, rowNum - 1, 1, width)
            af.LibxlDimLenToCell(book, sheet, rowNum - 1, 2, height)
            af.LibxlNumToCell(book, sheet, rowNum - 1, 3, scan - i1)
            af.LibxlCalcLenToCell(book, sheet, rowNum - 1, 4, totLen)
            af.LibxlAreaToCell(book, sheet, rowNum - 1, 5, totAreaGross)
            af.LibxlAreaToCell(book, sheet, rowNum - 1, 6, totAreaNet)
            af.LibxlVolToCell(book, sheet, rowNum - 1, 7, totVol)
            -- af.LibxlMbsToCell(book, sheet, rowNum-1, 8, colour)

            local col, coltype, colval

            -- dump(currplank.adjcols)
            for col = 0, ESummaryCols - 1 do
                coltype = gtblSummaryCols[col + 1]
                colval = currplank.adjcols[col + 1]
                -- af.Log(string.format("type=%s val=%s", tostring(coltype), tostring(colval)))
                -- coltype nil means no info selected, colval nil means (colour) value not present for the item
                if coltype and colval then
                    if coltype == ESummaryInfoColourSample then
                        local r, g, b, dummy
                        dummy, dummy, r, g, b =
                            string.find(colval, "([0123456789.]+),([0123456789.]+),([0123456789.]+)")
                        if r and g and b then
                            --- local color = book:color_pack(math.floor(255 * r + 0.5), math.floor(255 * g + 0.5), math.floor(255 * b + 0.5))
                            local color = af.LibxlColor(math.floor(255 * r + 0.5), math.floor(255 * g + 0.5),
                                math.floor(255 * b + 0.5))
                            local fmt = book:add_format()
                            if fmt ~= nil then
                                if gXlsx then
                                    color = book:color_pack(math.floor(255 * r + 0.5), math.floor(255 * g + 0.5),
                                        math.floor(255 * b + 0.5))
                                    fmt:set_border_color(color)
                                    fmt:set_fill_pattern(1) -- Solid
                                    fmt:set_pattern_fg_color(color)
                                else
                                    fmt:set_fill_pattern(1) -- Solid
                                    fmt:set_pattern_fg_color(color)
                                end
                                sheet:set_format(rowNum - 1, col, fmt)
                            end
                        end
                    elseif coltype ~= ESummaryInfoNone then
                        -- All these are texts
                        af.LibxlMbsToCell(book, sheet, rowNum - 1, col, tostring(colval))
                    end
                end
            end
        end

        rowNum = rowNum + 1
        i1 = scan
    end

    if othersNow then
        local tblSorted, tbl

        tblSorted = {}
        i1 = 0
        for i, v in pairs(othersNow) do
            tbl = {}
            tbl.key = i
            tbl.sortkey = i
            if v.sortkey then
                tbl.sortkey = v.sortkey
            end
            tbl.val = v -- has val.num and val.type
            i1 = i1 + 1
            tblSorted[i1] = tbl
        end

        table.sort(tblSorted, function(n1, n2)
            return n1.sortkey < n2.sortkey
        end)

        for i, v in ipairs(tblSorted) do
            -- Set some values, nil=missing
            local accname, count, len, vol

            accname = v.key
            if v.val.name then
                accname = v.val.name
            end

            if v.val.len then
                len = v.val.len
            end

            if v.val.type then
                if v.val.type == "len" then
                    len = v.val.num
                elseif v.val.type == "vol" then
                    vol = v.val.num
                end
            else
                count = v.val.num -- pcs
            end

            if sheet == nil then
                -- To txt file
                local sLine

                sLine = string.format("%s\t", accname)
                if v.val.width then
                    sLine = sLine .. ac_environment("ntos", v.val.width, "length", "dim")
                end
                sLine = sLine .. "\t"

                if v.val.height then
                    sLine = sLine .. ac_environment("ntos", v.val.height, "length", "dim")
                end
                sLine = sLine .. "\t"

                -- count
                if count then
                    sLine = sLine .. string.format("%.0f", count)
                end
                sLine = sLine .. "\t"

                -- len
                if len then
                    sLine = sLine .. ac_environment("ntos", len, "length", "calc")
                end
                sLine = sLine .. "\t"

                -- Extra fields
                if v.val.areagross then
                    sLine = sLine .. ac_environment("ntos", v.val.areagross, "area", "calc")
                end
                sLine = sLine .. "\t"

                if v.val.areanet then
                    sLine = sLine .. ac_environment("ntos", v.val.areanet, "area", "calc")
                end
                sLine = sLine .. "\t"

                -- vol
                if vol then
                    sLine = sLine .. ac_environment("ntos", vol, "volume", "calc") .. "\t"
                end
                sLine = sLine .. "\n"
                io.write(sLine)

            else
                -- To xls file
                af.LibxlMbsToCell(book, sheet, rowNum - 1, 0, accname)

                if v.val.width then
                    af.LibxlDimLenToCell(book, sheet, rowNum - 1, 1, v.val.width)
                end

                if v.val.height then
                    af.LibxlDimLenToCell(book, sheet, rowNum - 1, 2, v.val.height)
                end

                -- count
                if count then
                    af.LibxlNumToCell(book, sheet, rowNum - 1, 3, count)
                end

                -- len
                if len then
                    af.LibxlCalcLenToCell(book, sheet, rowNum - 1, 4, len)
                end

                -- Extra fields
                if v.val.areagross then
                    af.LibxlAreaToCell(book, sheet, rowNum - 1, 5, v.val.areagross)
                end

                if v.val.areanet then
                    af.LibxlAreaToCell(book, sheet, rowNum - 1, 6, v.val.areanet)
                end

                -- vol
                if vol then
                    af.LibxlVolToCell(book, sheet, rowNum - 1, 7, vol)
                end
            end
            rowNum = rowNum + 1
        end
    end

    return rowNum
end

function Planks2TxtSummary(tblOthers)
    local hFile

    hFile = io.open(gXlsName, "wt")
    if not hFile then
        error("Opening file " .. gXlsName .. " for writing failed. CHECK THAT FILE IS NOT OPEN IN EXCEL ETC.")
    end
    io.output(hFile)

    Planks2XlsSummary(nil, nil, tblOthers)

    io.close(hFile)
end

---------------------------------------------------------------------
-- Special listings element work

function OnInitElemWork(strPlnFileName)
    sFileName = strPlnFileName .. "_elemwork.xlsx"
    sExt = "xlsx"

    return sFileName, sExt
end

function OnSaveListElemWork(strFileName)
    local i, v, tblElemGuid2Master, tblElemGuid2Owner, tblPlanksData, sId, plankinfo, elemMaster, elemOwner, tblPlank,
        nPlanks, prevScript

    prevScript = gScriptUtf8
    gScriptUtf8 = 1

    -- Do the grouping by master element here, for performance have a helper table: element guid -> table master element info (guid, id)

    tblElemGuid2Master = {} -- key=element guid, value=tbl with fields: guid=master element guid, id=its id for sorting, type=type attribute from xml <layer ref="WALL 42x42 VERT" ... type="intstud">
    tblElemGuid2Owner = {} -- As previous, but only guid, type
    tblPlanksData = {}

    nPlanks = 0
    for i, v in ipairs(gTblPlanks) do
        tblPlank = {}

        ac_objectopen(v)
        plankinfo = af_request("plankinfo")
        tblPlank.guid = v
        tblPlank.id = ac_objectget("#id")

        -- Skip if not part of an element
        if plankinfo.ownerelemguid then
            elemMaster = tblElemGuid2Master[plankinfo.ownerelemguid]
            elemOwner = tblElemGuid2Owner[plankinfo.ownerelemguid]

            if elemMaster == nil or elemOwner == nil then
                -- Open the parent element
                local elemParent, iElem, vElem, elemMasterTemp

                elemParent = af_request("elem_openparent", plankinfo.ownerelemguid)
                elemMaster = tblElemGuid2Master[elemParent.guid]

                elemMasterTemp = {}
                elemMasterTemp.guid = elemParent.guid

                elemOwner = {}
                elemOwner.guid = plankinfo.ownerelemguid

                for iElem, vElem in ipairs(elemParent.tblelems) do
                    if vElem.guid == elemParent.guid then
                        elemMasterTemp.type = vElem.type
                    end
                    if vElem.guid == plankinfo.ownerelemguid then
                        elemOwner.type = vElem.type
                    end
                end

                if elemMasterTemp.type == nil or elemOwner == nil then
                    af.RaiseError(string.format("Cannot find parent element for plank %s", v))
                end

                ac_objectclose()
                ac_objectopen(elemMasterTemp.guid)
                elemMasterTemp.id = ac_objectget("#id")
                elemMasterTemp.floor = ac_objectget("#floor")
                ac_objectclose()
                ac_objectopen(v)

                if elemMaster == nil then
                    elemMaster = elemMasterTemp
                end

                tblElemGuid2Master[elemParent.guid] = elemMaster
                tblElemGuid2Master[plankinfo.ownerelemguid] = elemMaster
                tblElemGuid2Owner[plankinfo.ownerelemguid] = elemOwner
            end

            nPlanks = nPlanks + 1
            tblPlank.index = nPlanks
            tblPlank.tblMaster = elemMaster
            tblPlank.tblOwnerElem = elemOwner
            SetElemGroup(tblPlank)

            tblPlanksData[nPlanks] = tblPlank
        end
        ac_objectclose()
    end

    -- Framen exceli
    gXlsName = strFileName
    gtblPlanksData = tblPlanksData
    local status, err = pcall(DoFrameExcelElemWork)

    gScriptUtf8 = prevScript

    if not status then
        if excelStarted ~= nil then
            ws = nil
            wb = nil
            excelStarted:Quit()
            excelStarted = nil
        end
        af.RaiseError("Creating listing failed: " .. err)
    end
end

-- Compares:
-- 1. Master element id
-- 2. Grouping sort key
-- 3. Plank id
-- Returns: -1=n1<n2, 0=same, 1=n1>n2
function CmpElemWorkPlank(n1, n2)
    -- Element id
    if n1.tblMaster.id < n2.tblMaster.id then
        return -1
    elseif n1.tblMaster.id > n2.tblMaster.id then
        return 1
    end

    -- Master element guid, nope, group just by visible id
    -- if n1.tblMaster.guid<n2.tblMaster.guid then
    --	return true
    -- elseif n1.tblMaster.guid>n2.tblMaster.guid then
    --	return false
    -- end

    -- Grouping text
    if n1.elemgroupsort < n2.elemgroupsort then
        return -1
    elseif n1.elemgroupsort > n2.elemgroupsort then
        return 1
    end

    return 0
end

function DoFrameExcelElemWork()
    local templateName, wb
    local str, file, fileExt
    local projNumber, useNumToStr, res

    -- Sort by master element guid and plank index
    table.sort(gtblPlanksData, function(n1, n2)
        local cmp

        cmp = CmpElemPlank(n1, n2)
        if cmp ~= 0 then
            return cmp < 0
        end

        return n1.index < n2.index -- Keep the order if at same pos (id is part of sort order)
    end)

    -- Do we have imperial working settings?
    af.CheckImperial()

    templateName = XlsxGetTemplateFileName("FrameListingElemWork", ".xlsx")
    fileExt = af.GetFileExt(gXlsName)
    if fileExt ~= ".xls" and fileExt ~= ".xlsx" then
        Planks2TxtElemWork()
        return
    end

    -- Create workbook
    local book

    book = af.LibxlCreateBook(fileExt, templateName)

    Planks2XlsElemWork(book)

    if book:save(gXlsName) == false then
        af.RaiseError("Failed to save Excel workbook `" .. gXlsName ..
                          "`\nPlease make sure the file is not open in Excel already.")
    end

    book:release()

    -- Open the newly created one in excel
    af.ExcelOpen(gXlsName)
end

function ReadAcc(ws, row, col)
    local s

    s = ws:read_str(row, col)
    if s and s == "" then
        s = nil
    end
    return s
end

-- Saves txt file if wb is null (mac or text output anyway)
function Planks2XlsElemWork(wb)
    local i1, i2, tblPlanks, nPlanks, ws, scan
    local tblFloors = {} -- Key=floor index 1-N, value=floor name
    local tblElems = {} -- Key=element index 1-N, value=table having fields: id=element ID which is same as worksheet name, floor=floor name
    local tblAccNames = {}

    -- Collect names used in ArchiFrameAccessory-table object. field name rowx is zero based
    ws = wb:get_sheet(1)
    if ws == nil then
        error("Definition worksheet is missing")
    end
    tblAccNames.row4_element = ReadAcc(ws, 4, 5)
    tblAccNames.row7_prefab = ReadAcc(ws, 7, 5)
    tblAccNames.row9_gn = ReadAcc(ws, 9, 5)
    tblAccNames.row10_gf = ReadAcc(ws, 10, 5)
    tblAccNames.row11_ws = ReadAcc(ws, 11, 5)
    tblAccNames.row12_aqua = ReadAcc(ws, 12, 5)
    tblAccNames.row13_nails = ReadAcc(ws, 13, 5)
    tblAccNames.row16_vapour = ReadAcc(ws, 16, 5)
    tblAccNames.row17_claddingextra = ReadAcc(ws, 16, 5)
    tblAccNames.row21_clad = ReadAcc(ws, 21, 5)
    tblAccNames.row22_cladhole = ReadAcc(ws, 22, 5)
    tblAccNames.row23_screw = ReadAcc(ws, 23, 5)
    tblAccNames.row25_mep = ReadAcc(ws, 25, 5)
    tblAccNames.row26_iron = ReadAcc(ws, 26, 5)
    tblAccNames.row27_balkshoe = ReadAcc(ws, 27, 5)
    tblAccNames.row28_plastic = ReadAcc(ws, 28, 5)
    tblAccNames.row29_rodent = ReadAcc(ws, 29, 5)
    tblAccNames.row30_transport = ReadAcc(ws, 30, 5)
    tblAccNames.row31_lift = ReadAcc(ws, 31, 5)

    -- Process per element
    i1 = 1
    while gtblPlanksData[i1] do
        -- Create table for planks in single elem
        i2 = i1
        nPlanks = 0
        tblPlanks = {}
        while true do
            if gtblPlanksData[i2] == nil then
                break
            end
            if gtblPlanksData[i1].tblMaster.id ~= gtblPlanksData[i2].tblMaster.id then
                break
            end

            nPlanks = nPlanks + 1
            tblPlanks[nPlanks] = gtblPlanksData[i2]
            i2 = i2 + 1
        end

        -- af.Log( string.format("i1=%d i2=%d", i1, i2) )
        -- dump(gtblPlanksData[i1].tblMaster)
        -- af.Log("")

        local floorName = ac_getobjparam(gtblPlanksData[i1].tblMaster.guid, "#floor")

        -- Add floor name if new
        scan = 1
        while scan <= #tblFloors do
            if floorName == tblFloors[scan] then
                break
            end
            scan = scan + 1
        end
        tblFloors[scan] = floorName
        -- af.Log(string.format("floor=%s elemid=%s scan=%d", floorName, gtblPlanksData[i1].tblMaster.id, scan))

        -- Every element ID should be different - add elem id
        local t = {}

        t.id = gtblPlanksData[i1].tblMaster.id
        t.floor = floorName
        tblElems[#tblElems + 1] = t

        Planks2XlsElemWorkInt(gtblPlanksData[i1].tblMaster, wb, tblPlanks, tblAccNames)
        i1 = i2
    end

    -- # Delete template
    local count

    count = wb:sheet_count()
    wb:del_sheet(count - 1)

    -- Show summary
    wb:activate_sheet(0)

    -- # Add summary page, sort floors
    table.sort(tblFloors, function(s1, s2)
        return s1 < s2
    end)

    -- Go through floors
    local kf, vf, ke, ve, ws

    ws = wb:get_sheet(0)

    projName = ac_environment("parsetext", "<PROJECTNAME>")
    if projName == nil then
        projName = "?"
    end
    af.LibxlMbsToCell(wb, ws, 0, 0, projName)
    local kolmas, row, col, eikolmas, sGrossm2Tot, sNetm2Tot, sHoursTot, sHoursPerM2, nElems
    local sFloor

    kolmas = 3
    eikolmas = 0
    row = 3
    col = 3
    sGrossm2Tot = ""
    sNetm2Tot = ""
    sHoursTot = ""
    sHoursPerM2 = ""
    nElems = 0

    local sBrm2, sNetM2, sHtot, sHbrM2

    sBrM2 = "grsm²"
    sNetM2 = "netm²"
    sHtot = "total hours"
    sHbrM2 = "h/brtm²"
    sTOT = "Total"
    sFloor = "%s STOREY"
    if af.GetLangStr3() == "Fin" then
        sBrM2 = "brtm²"
        sNetM2 = "nettom²"
        sHtot = "kokonaiskesto (h)"
        sHbrM2 = "h/brtm²"
        sTOT = "Kokonaisuus"
        sFloor = "%s KERROS"
    end

    for kf, vf in ipairs(tblFloors) do
        -- Find all elements belonging to this floor (vf)
        local tblElemsNow = {}

        kolmas = 3 -- nollataan
        col = 3 -- palautetaan rivin ekaksi

        -- af.Log(string.format("FLOOR=%s", vf))
        af.LibxlMbsToCell(wb, ws, row - 1, 0, string.format(sFloor, vf))

        row = row + 1

        for ke, ve in ipairs(tblElems) do
            if ve.floor == vf then
                tblElemsNow[#tblElemsNow + 1] = ve
            end
        end

        -- Sort elements by ID
        table.sort(tblElemsNow, function(s1, s2)
            return s1.id < s2.id
        end)

        -- Process each element (must be at least now that we are here)
        for ke, ve in ipairs(tblElemsNow) do
            -- af.Log(string.format(" elem id=%s", ve.id))
            --	ws.Cells(row  ,col).Value2 = string.format("=\'%s\'!A1", ve.id)
            --	ws.Cells(row+1,col).Value2 = string.format("=\'%s\'!B1", ve.id)
            --	ws.Cells(row+2,col).Value2 = string.format("=\'%s\'!B2", ve.id)
            --	ws.Cells(row+3,col).Value2 = string.format("=\'%s\'!E29", ve.id)
            --	ws.Cells(row+4,col).Value2 = string.format("=\'%s\'!E30", ve.id)

            local vId, vBr, vNe, vHt, vHb
            vId = string.format("\'%s\'!A1", ve.id)
            vBr = string.format("\'%s\'!B1", ve.id)
            vNe = string.format("\'%s\'!B2", ve.id)
            vHt = string.format("\'%s\'!E33", ve.id)
            vHb = string.format("\'%s\'!E34", ve.id)

            ws:write_formula(row - 1, col - 1, vId)
            ws:write_formula(row + 0, col - 1, vBr)
            ws:write_formula(row + 1, col - 1, vNe)
            ws:write_formula(row + 2, col - 1, vHt)
            ws:write_formula(row + 3, col - 1, vHb)

            if sGrossm2Tot ~= "" then
                sGrossm2Tot = sGrossm2Tot .. "+"
                sNetm2Tot = sNetm2Tot .. "+"
                sHoursTot = sHoursTot .. "+"
                sHoursPerM2 = sHoursPerM2 .. "+"
            end
            sGrossm2Tot = sGrossm2Tot .. string.format("%s%d", GetColChar(col), row + 1) -- Refer to summary page's cell
            sNetm2Tot = sNetm2Tot .. string.format("%s%d", GetColChar(col), row + 2)
            sHoursTot = sHoursTot .. string.format("%s%d", GetColChar(col), row + 3)
            sHoursPerM2 = sHoursPerM2 .. string.format("%s%d", GetColChar(col), row + 4)
            nElems = nElems + 1

            af.LibxlMbsToCell(wb, ws, row + 0, col, sBrM2)
            af.LibxlMbsToCell(wb, ws, row + 1, col, sNetM2)
            af.LibxlMbsToCell(wb, ws, row + 2, col, sHtot)
            af.LibxlMbsToCell(wb, ws, row + 3, col, sHbrM2)

            if ke == kolmas then
                kolmas = kolmas + 3
                row = row + 6
                col = 3
                eikolmas = 0
            else
                row = row
                col = col + 4
                eikolmas = 1
            end
        end
        if eikolmas > 0 then
            row = row + 6
        end
        col = 3

    end

    af.LibxlMbsToCell(wb, ws, row - 1, 0, sTOT)
    ws:write_formula(row + 0, col - 1, sGrossm2Tot)
    ws:write_formula(row + 1, col - 1, sNetm2Tot)
    ws:write_formula(row + 2, col - 1, sHoursTot)

    if nElems == 0 then
        nElems = 1
    end
    ws:write_formula(row + 3, col - 1, "(" .. sHoursPerM2 .. string.format(") / %d", nElems))

    af.LibxlMbsToCell(wb, ws, row + 0, col, sBrM2)
    af.LibxlMbsToCell(wb, ws, row + 1, col, sNetM2)
    af.LibxlMbsToCell(wb, ws, row + 2, col, sHtot)
    af.LibxlMbsToCell(wb, ws, row + 3, col, sHbrM2)
end

-- Removes all hidden objects and builds result table having fields: guid, libguid
function FilterAccTbl(tblAcc)
    if tblAcc == nil then
        return nil
    end

    local t, res
    local k, v

    res = {}
    for k, v in ipairs(tblAcc) do
        if ac_environment("filterelem", v, 2) then
            t = {}
            t.guid = v
            t.libguid = ac_elemget(v).header.libGuid
            res[#res + 1] = t
        end
    end

    return res
end

-- Processes visible items
-- Returns two values:
-- 1: Number of electric boxes+
-- 2: Table of collected lines from found ArchiFrameAccessory-objects: key=name, data fields: num, unit
function CalcAccTbl(tblAcc)
    local k, v, nele, s, t
    local accItems = {}

    nele = 0
    for k, v in ipairs(tblAcc) do

        s = string.match(v.libguid, "{(.-)}") -- GUID is in form {22CFB950-C5CA-49DC-9C8F-16EBB3012153}-{CD05BACD-764E-466A-9CFE-49F5F7BAFB1B}, take first
        -- af.Log(string.format("libguid=%s s=%s", v.libguid, s))
        if s == "22CFB950-C5CA-49DC-9C8F-16EBB3012153" then
            -- El box
            ac_objectopen(v.guid)
            nele = nele + ac_objectget("iRows") * ac_objectget("iCols") -- + ac_objectget("iWiresOut")
            ac_objectclose()
        elseif s == "63C2BF51-7EDD-441D-B7ED-DF7E7832BE52" then
            local i, rows, name, num, unit

            ac_objectopen(v.guid)
            rows = ac_objectget("iNames", -1)
            for i = 1, rows do
                name = ac_objectget("iNames", i)
                num = ac_objectget("iAmounts", i)
                unit = ac_objectget("iUnits", i)
                t = accItems[name]
                if t == nil then
                    t = {}
                    t.num = num
                    t.unit = unit
                    accItems[name] = t
                else
                    t.num = t.num + num
                end
            end
            ac_objectclose()
        end
    end
    return nele, accItems
end

function MatchAccRow(str, wildstr)
    local word

    str = string.lower(str)
    wildstr = string.lower(wildstr)

    for word in string.gmatch(wildstr, '([^,]+)') do
        -- af.Log(string.format("word=%s", word))
        if string.match(str, word) then
            return true
        end
    end
    return false
end

-- keywild can contain Lua wildcards
-- Returns num of requested acc item. If it does not exist, issues warning
function GetAccItem(elemid, tblAccItems, keywild)
    local k, v

    if not keywild then
        return 0
    end

    if tblAccItems then -- Cannot use for string keys: and #tblAccItems>0 
        local num, found, data

        -- Scan all items to support wild cards
        num = 0
        for k, v in pairs(tblAccItems) do
            if MatchAccRow(k, keywild) then
                found = true
                num = num + v.num
            end
            data = true
        end

        if found then
            return num
        end
        if data then
            -- Was it empty?
            af.Log(string.format(
                "WARNING: Element %s: ArchiFrameAccessory-object is missing setting %s. Using default value 0", elemid,
                keywild))
        end
    end

    return 0
end

-- Saves txt file if wb is null (mac or text output anyway)
-- Writes planks for single element
-- rowNum		Number of first row to write
-- tblPlanks	Ignored here - always calculates the whole element
-- Returns next unused row number
function Planks2XlsElemWorkInt(tblMaster, wb, tblPlanks, tblAccNames)
    local colIds, colMatId, colLength, colAngleBeg, colAngleEnd, colCount, prevElemgroup, colElemGroup
    local cell
    local matLen, matInfo, val, i, v, name, strLen
    local idNow, matId, height, width, len, usage, similarCount, idList, totalCount, col
    local i, scan
    local tblCells = {} -- Cells for single row (to be saved to xls/txt)
    local tblCellFormatTxt = {} -- printf format strings if saving to txt file
    local tblQuant -- For all layers

    -- Kaikki vakiotekstit näin (UNUSED FOR XLS)
    local sGrossM2, sNetM2, sPcs, sM2, sM

    sGrossM2 = "grossm²"
    sNetM2 = "netm²"
    sPcs = "pcs"
    sM2 = "m²"
    sM = "m"

    if af.GetLangStr3() == "Fin" then
        sGrossM2 = "brtm²"
        sNetM2 = "nettom²"
        sPcs = "kpl"
        sM2 = "m²"
        sM = "jm"
    end

    tblQuant = af_request("elem_quantities", tblMaster.guid, 1)
    -- dump(tblQuant)

    -- COMMON CALCULATION LOGIC FOR XLS AND TXT
    local k, v, nPlanks, nPlanks25, t, meterBeams, paino

    meterBeams = 0
    nPlanks = 0
    nPlanks25 = 0
    paino = 0

    local nElem, vElem, vElemCore
    local claddinghorm2, claddingvertm2, koolaus, claddingholesm2
    local gnm2, gfm2, gtsm2, aquam2
    -- local 	villa100m2
    local insuMain, insuStud -- For main framing and studding layers
    local coreGrossM2, coreNetM2
    local planksMul
    local tblAcc = {} -- CalcMcQuant() calcs here

    claddinghorm2 = 0
    claddingvertm2 = 0
    koolaus = 0
    claddingholesm2 = 0

    gnm2 = 0
    gfm2 = 0
    gtsm2 = 0
    aquam2 = 0
    insuMain = 0
    insuStud = 0
    coreGrossM2 = 0
    coreNetM2 = 0

    for nElem, vElem in ipairs(tblQuant.tblelems) do
        local extCladding = false

        -- af.Log(string.format("vElem.elemtype=%s", vElem.elemtype))
        planksMul = 1
        if vElem.elemtype == "core" then
            vElemCore = vElem
            coreGrossM2 = vElem.areagross
            coreNetM2 = vElem.areanet
            if vElem.elemweightreduced > 0 then
                -- Insulation with core, add net area
                insuMain = coreNetM2
            end
        elseif vElem.elemtype == "extstud" then
            if vElem.tblplanks then
                koolaus = koolaus + #vElem.tblplanks
            end
            planksMul = 0
            if vElem.elemweightreduced > 0 then
                -- Insulation with studding, add net area (allow also to have with extstud)
                insuStud = insuStud + vElem.areanet
            end
        elseif vElem.elemtype == "intstud" then
            -- Added to total amount: koolaus=koolaus+#vElem.tblplanks
            if vElem.elemweightreduced > 0 then
                -- Insulation with studding, add net area
                insuStud = insuStud + vElem.areanet
            end
        elseif vElem.elemtype == "finish_ext" then
            -- Calculate from related boardpanel. NOTE! If exploded, will not return good value.
            planksMul = 0
            if vElem.tblplanks then
                af.Log(string.format(
                    "WARNING: Element %s has exploded cladding - calculating cladding area hole from related ArchiFrameBoardPanel-objects only",
                    tblMaster.id))
            end

            local tblSettings, tblPolys1, tblPolys2

            -- First area without holes
            tblSettings = {}
            tblSettings.fromboard = 1
            tblSettings.holes = 0
            tblPolys1 = af_request("getpoly", tblSettings, vElem.guid)
            if tblPolys1 == nil or tblPolys1.area == 0 then
                tblSettings.fromboard = nil
                tblPoly1s = af_request("getpoly", tblSettings, vElem.guid)
                af.Log(string.format(
                    "WARNING: Element %s does not have related ArchiFrameBoardPanel-objects - calculating cladding area from related ArchiFrameElement-object",
                    tblMaster.id))
            end

            -- Then area with holes
            tblSettings.holes = nil
            tblPolys2 = af_request("getpoly", tblSettings, vElem.guid)
            -- af.Log(string.format("area1=%f area2=%f", tblPolys1.area, tblPolys2.area))
            claddingholesm2 = claddingholesm2 + tblPolys1.area - tblPolys2.area

            local rot

            -- Classify whole layer to vertical/horizontal from the first related board
            rot = 0
            if vElem.tblboards and vElem.tblboards[1] then
                rot = ac_getobjparam(vElem.tblboards[1].guid, "iPanelRot")
            elseif vElem.tblplanks and vElem.tblplanks[1] then
                -- Exploded - take direction from the first plank
                local info

                info = af_request("plankinfo", vElem.tblplanks[1].guid)
                if math.abs(info.vecx.z) > 0.7 then
                    rot = PI / 2
                end
            end

            if math.abs(rot) > PI / 4 then
                claddingvertm2 = claddingvertm2 + tblPolys1.area
            else
                claddinghorm2 = claddinghorm2 + tblPolys1.area
            end

            extCladding = true
        end

        -- Skip exploded planks by layer type
        if vElem.tblplanks and vElem.elemtype ~= "finish_ext" and vElem.elemtype ~= "finish_int" then
            local tblSettings

            tblSettings = {}
            tblSettings.quant = 1 -- We need plank weights

            for k, v in pairs(vElem.tblplanks) do
                ac_objectopen(v.guid)
                t = af_request("plankinfo", nil, nil, tblSettings)
                -- Plank
                group = ac_objectget("iElemGroup")
                group = string.lower(group)
                if string.match(group, "^balk.*") then
                    if t.width > 0.0095 then -- CHECK: DO NOT CALCULATE BOARDS
                        meterBeams = meterBeams + t.len
                    end
                end

                -- Plank weight
                if t.kgpermeter then
                    paino = t.len * t.kgpermeter
                    -- af.Log(string.format("%s: %f kg", ac_objectget("#id"), paino))
                else
                    if not gbVerWarned then
                        gbVerWarned = true
                        af.Log(
                            "WARNING: Program version does not give plank weight - estimating it here. Please update program.")
                    end
                    paino = (t.len * t.width * t.height) * 0.45 * 1000
                end

                if paino > 24.999 then
                    nPlanks25 = nPlanks25 + 1 * planksMul
                else
                    nPlanks = nPlanks + 1 * planksMul
                end

                CalcMcQuant(tblAcc) -- For balk shoes and other accessories ****************************************************************************

                ac_objectclose()
            end
        end

        if vElem.tblboards then
            for k, v in pairs(vElem.tblboards) do
                local areagross, matid, rot, areanet

                ac_objectopen(v.guid)
                areanet = ac_objectget("iArea")
                t = af_request("plankinfo")
                areagross = t.len * t.height

                matid = ac_objectget("iMatId")
                matid = string.lower(matid)
                rot = ac_objectget("iPanelRot")

                if v.ispanel == 1 then
                    if extCladding then
                        -- Take polygon area without reducing holes, old code
                        --[[
						if math.abs(rot)>PI/4 then
							claddingvertm2=claddingvertm2+areanet
						else
							claddinghorm2=claddinghorm2+areanet
						end
						]]
                        -- Use area calculated earlier
                    end

                elseif MatchAccRow(matid, tblAccNames.row12_aqua) then
                    aquam2 = aquam2 + areagross
                elseif MatchAccRow(matid, tblAccNames.row11_ws) then
                    gtsm2 = gtsm2 + areagross
                elseif MatchAccRow(matid, tblAccNames.row10_gf) then
                    gfm2 = gfm2 + areagross
                elseif MatchAccRow(matid, tblAccNames.row9_gn) then
                    gnm2 = gnm2 + areagross
                else
                    gnm2 = gnm2 + areagross
                    af.Log(string.format("Unknown board: %s included to gypsum into line 10", matid))
                end

                ac_objectclose()
            end
        end
    end

    -- All ArchiFrameElMarker and ArchiFrameAccessory-objects
    local tblAcc = af_request("mc_getacc", tblMaster.guid)
    local nElBoxes = 0
    local tblAccItems -- Table of collected lines from found ArchiFrameAccessory-objects: key=name, data fields: num, unit

    tblAcc = FilterAccTbl(tblAcc)
    if tblAcc then
        nElBoxes, tblAccItems = CalcAccTbl(tblAcc)

        --[[
		local k,v

		for k,v in pairs(tblAccItems) do
			af.Log(string.format("name=%s num=%f unit=%s", k, v.num, v.unit))
		end
		--]]
    end

    if wb then
        -- Save to Excel
        local ws, count, num

        -- New sheet
        count = wb:sheet_count()
        ws = wb:insert_sheet(count - 1, count - 1, tblMaster.id)
        ws = wb:get_sheet(count - 1)

        rowNum = 0

        af.LibxlMbsToCell(wb, ws, rowNum, 0, tblMaster.id)
        af.LibxlAreaToCell(wb, ws, rowNum, 1, vElemCore.areagross)
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, vElemCore.areanet)
        rowNum = rowNum + 3

        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row4_element)) -- Elem starting time
        rowNum = rowNum + 1

        af.LibxlNumToCell(wb, ws, rowNum, 1, nPlanks)
        rowNum = rowNum + 1

        af.LibxlNumToCell(wb, ws, rowNum, 1, nPlanks25)
        rowNum = rowNum + 1

        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row7_prefab))
        rowNum = rowNum + 1

        af.LibxlCalcLenToCell(wb, ws, rowNum, 1, meterBeams)
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, gnm2)
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, gfm2)
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, gtsm2)
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, aquam2)
        rowNum = rowNum + 1

        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row13_nails))
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, insuStud)
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, insuMain)
        rowNum = rowNum + 1

        num = GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row16_vapour) -- Let it be multiplication factor
        af.LibxlAreaToCell(wb, ws, rowNum, 1, num * coreGrossM2) -- Vapor plastic M2
        rowNum = rowNum + 1

        num = GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row17_claddingextra)
        if num ~= 0 then
            num = 1
        end
        af.LibxlAreaToCell(wb, ws, rowNum, 1, num * coreNetM2) -- Korokepala, TODO
        rowNum = rowNum + 1

        af.LibxlNumToCell(wb, ws, rowNum, 1, koolaus)
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, claddinghorm2)
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1, claddingvertm2)
        rowNum = rowNum + 1

        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row21_clad))
        rowNum = rowNum + 1

        af.LibxlAreaToCell(wb, ws, rowNum, 1,
            claddingholesm2 + GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row22_cladhole))
        rowNum = rowNum + 1

        -- Tarvikkeita
        -- Kansiruuvi
        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row23_screw))
        rowNum = rowNum + 1

        -- Sähkö sis. Rasiapohjan
        af.LibxlNumToCell(wb, ws, rowNum, 1, nElBoxes)
        rowNum = rowNum + 1

        -- IV
        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row25_mep))
        rowNum = rowNum + 1

        -- Kulmaraudat
        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row26_iron))
        rowNum = rowNum + 1

        -- Palkkikengät
        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row27_balkshoe))
        rowNum = rowNum + 1

        -- Suojamuovi
        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row28_plastic))
        rowNum = rowNum + 1

        -- Pieneläinverkko
        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row29_rodent))
        rowNum = rowNum + 1

        -- Kuljetustuet
        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row30_transport))
        rowNum = rowNum + 1

        -- Nostovahvikkeet
        af.LibxlNumToCell(wb, ws, rowNum, 1, GetAccItem(tblMaster.id, tblAccItems, tblAccNames.row31_lift))
        rowNum = rowNum + 1
    else
        -- TXT
    end

    return rowNum
end

function Planks2TxtElem()
    local hFile

    hFile = io.open(gXlsName, "wt")
    if not hFile then
        error("Opening file " .. gXlsName .. " for writing failed. CHECK THAT FILE IS NOT OPEN IN EXCEL ETC.")
    end
    io.output(hFile)

    Planks2XlsElem(nil)

    io.close(hFile)
end




























-- ARCHIFRAME CONTROL CARD LISTING (Readable Format)
--------------------------------------------------------------------------------
-- Utility: Get layer name by GUID
--------------------------------------------------------------------------------
function GetLayerNameOfObj(guid)
    local elem = ac_elemget(guid)
    if not elem or not elem.header then
        return ""
    end
    local attr = ac_getattrinfo(2, elem.header.layer)
    return (attr and attr.name) or ""
end

--------------------------------------------------------------------------------
-- Save-as Dialog
--------------------------------------------------------------------------------
function OnInitControlCard(filename)
    -- Returns new filename with suffix and xlsm extension
    return filename .. "_ControlCard.xlsm", "xlsm"
end

--------------------------------------------------------------------------------
-- Control Card Exporter (Aggregated + Colour Debug)
--------------------------------------------------------------------------------
function OnSaveListControlCard(outFile)
    -- Load Excel template
    local ext = GetFileExt(outFile)
    local tpl = XlsxGetTemplateFileName("ControlCardTemplate", ext) -- original template name -- use Eng template

    local fh = io.open(tpl, "rb")
    if not fh then
        af.RaiseError("Template not found: " .. tpl)
    end
    fh:close()

    local book = af.LibxlCreateBook(ext, tpl)
    local ws   = book:get_sheet(0)
    if not ws then
        af.RaiseError("Sheet 0 not found in template")
    end

    -- Header
    af.LibxlMbsToCell(book, ws, 1, 2, GetAutoTextNoNil("<PROJECTNAME>"))
    af.LibxlMbsToCell(book, ws, 2, 2, GetAutoTextNoNil("<PROJECTNUMBER>"))
    af.LibxlMbsToCell(book, ws, 3, 2, GetAutoTextNoNil("<CLIENT>"))
    af.LibxlMbsToCell(book, ws, 4, 2, GetAutoTextNoNil("<DESIGNER>"))

    -- Aggregate data
    local agg = {}

    for _, guid in ipairs(gTblPlanks) do
        ac_objectopen(guid)

        -- Common properties
        local id     = ac_objectget("#id")    or ""
        local layer  = GetLayerNameOfObj(guid)
        local mat    = ac_objectget("iMatId") or ""

        -- Default board dimensions
        local w      = ac_objectget("iWidth")  or 0
        local h      = ac_objectget("iHeight") or 0
        local len    = af.GetPlankLength()      or 0
        local len_mm = ac_environment("ntos", len, "length", "dim") or ""

        -- Actual panel dimensions if this is a panel (thickness <=10mm)
        local panel_w, panel_h, panel_len
        if w <= 10 or h <= 10 then
            local info = af_request("plankinfo") or {}
            panel_w   = ac_objectget("iCurrWidth") or 0
            panel_h   = ac_objectget("iCurrHeight") or 0
            panel_len = info.width or len
        end

        -- Colour / Coating Name
        local col = ac_getobjparam(guid, "#matname") or ""
        if col == "" then
            local sm   = af_request("singlemat", mat)
            local coat = sm and string.match(sm.xmlutf8 or "", 'coating="(.-)"') or ""
            local exT  = sm and string.match(sm.xmlutf8 or "", 'excel_type="(.-)"') or ""
            col = (coat ~= "" and coat) or exT
        end
        if col == "" then
            local surfGuid = ac_objectget("#material") or ""
            if surfGuid ~= "" then
                local si = af_request("surfaceinfo", surfGuid)
                if si and si.name and si.name ~= "" then
                    col = si.name
                end
            end
        end

        -- Area and Volume Calculations
        local calc_w, calc_h, calc_len = w, h, len
        if panel_w and panel_w > 0 and panel_h and panel_h > 0 and panel_len and panel_len > 0 then
            calc_w   = panel_w
            calc_h   = panel_h
            calc_len = panel_len
        end

        local area1 = calc_w * (calc_len)       -- wide face m²
        local area2 = calc_h * (calc_len)       -- narrow face m²
        local vol   = calc_w * calc_h * (calc_len)   -- volume m³

        ac_objectclose()

        -- Aggregation key (includes colour to separate different coatings)
        local key = table.concat({id, w, h, len_mm, mat, layer, col}, "|")
        local rec = agg[key]
        if not rec then
            rec = {
                id     = id,
                layer  = layer,
                w      = w,
                h      = h,
                len    = len,
                len_mm = len_mm,
                mat    = mat,
                col    = col,
                area1  = area1,
                area2  = area2,
                vol    = vol,
                qty    = 0,
                panel_w   = panel_w, 
                panel_h   = panel_h,
                panel_len = panel_len,
            }
            agg[key] = rec
        end
        rec.qty = rec.qty + 1
    end

    -- Write to Excel
    -- Сортировка по площади сечения (используем panel dimensions, если есть)
    local sorted = {}
    for _, r in pairs(agg) do
        table.insert(sorted, r)
    end
    table.sort(sorted, function(a, b)
        local aW = (a.panel_w and a.panel_w > 0) and a.panel_w or a.w
        local aH = (a.panel_h and a.panel_h > 0) and a.panel_h or a.h
        local bW = (b.panel_w and b.panel_w > 0) and b.panel_w or b.w
        local bH = (b.panel_h and b.panel_h > 0) and b.panel_h or b.h
        return (aW * aH) < (bW * bH)
    end)

    local row = 11
    local num = 1
    for _, r in ipairs(sorted) do
        af.LibxlNumToCell(book, ws, row, 0, num)                        -- № Элем
        af.LibxlMbsToCell(book, ws, row, 1, r.id)                       -- ID
        af.LibxlMbsToCell(book, ws, row, 2, r.layer)                    -- Слой

                -- Сечение, мм: для плит (толщина < 10 мм) – фактические размеры детали, иначе – стандартное сечение
        if (r.len_mm*1) < 10 then
            -- плита, ширина (толщина) r.w : длина × высота
            af.LibxlMbsToCell(book, ws, row, 3,
                string.format("%sx%s",
                    ac_environment("ntos", r.panel_w, "length", "dim"),
                    ac_environment("ntos", r.panel_h,   "length", "dim"))) -- Сечение, мм (плита)
        else
            -- доска: толщина (r.w) × высота (r.h)
            af.LibxlMbsToCell(book, ws, row, 3,
                string.format("%sx%s",
                    ac_environment("ntos", r.w, "length", "dim"),
                    ac_environment("ntos", r.h, "length", "dim"))) -- Сечение, мм (доска)
        end

        af.LibxlMbsToCell(book, ws, row, 4, r.len_mm)                   -- Длина, мм
                -- Размер заготовки: получить из определения материала
        local matDef = gtblFrameMat[r.mat]
        local stockW = (matDef and matDef.width)  or r.w
        local stockH = (matDef and matDef.height) or r.h
        af.LibxlMbsToCell(book, ws, row, 9,
            string.format("%sx%s",
                ac_environment("ntos", stockW, "length", "dim"),
                ac_environment("ntos", stockH, "length", "dim")))   -- Длина, мм
        af.LibxlNumToCell(book, ws, row, 5, r.qty)                  -- Кол-во
        af.LibxlNumToCell(book, ws, row, 6, r.vol)                  -- Объём, м³
        af.LibxlNumToCell(book, ws, row, 7, r.area1)                -- Площадь грани №1
        af.LibxlNumToCell(book, ws, row, 8, r.area2)                -- Площадь грани №2
        af.LibxlMbsToCell(book, ws, row, 10, r.col)                 -- Цвет
        af.LibxlMbsToCell(book, ws, row, 11, r.mat)                 -- Материал
        row = row + 1
        num = num + 1
    end

    -- ===================================================
    -- === Запись в лист "Невидимки" ArchiFrameElement ===
    -- ===================================================
   local wsInv = book:get_sheet(1)  -- второй лист (индексация с 0)
    if not wsInv then
        af.RaiseError("Sheet 'Невидимки' not found in template!")
    end

    local nrow = 11 -- стартовая строка (12-я, индексация с 0)
    local n = 1
   
    -- Собираем множество guid из основной таблицы
    local gTblPlanksSet = {}
    for _, guid in ipairs(gTblPlanks) do
        gTblPlanksSet[guid] = true
    end

    function PolyLength(pts)
        local L = 0
        local count = #pts
        if count < 2 then return 0 end
        for i = 1, count do
            local j = (i % count) + 1 -- следующий, с замыканием
            local dx = (pts[j].x or 0) - (pts[i].x or 0)
            local dy = (pts[j].y or 0) - (pts[i].y or 0)
            L = L + math.sqrt(dx*dx + dy*dy)
        end
        return L
    end

    local function log_warning(msg)
        local log_file = io.open("ControlCard_Warnings.log", "a")
        if log_file then
            log_file:write(msg .. "\n")
            log_file:close()
        end
    end

    local selGuids = ac_environment("getsel") or {}
    for _, guid in ipairs(selGuids) do
        if not gTblPlanksSet[guid] then
            -- Защищённый вызов af_request
            local ok, q = pcall(af_request, "elem_quantities", guid, 0)
            if not ok or not q then
                log_warning("Ошибка или неверный тип для guid: " .. tostring(guid) .. " при elem_quantities")
            else
                local id = ""
                local ok_info, objinfo = pcall(af_request, "objectinfo", guid)
                if ok_info and objinfo and objinfo.id then
                    id = objinfo.id
                else
                    -- fallback: try reading ID directly from object
                    local opened = pcall(ac_objectopen, guid)
                    if opened then
                        id = ac_objectget("#id") or ""
                        ac_objectclose()
                    end
                    if id == "" then
                        log_warning("Ошибка или нет id для guid: " .. tostring(guid) .. " при objectinfo")
                    end
                end

                local ok_layer, layer = pcall(GetLayerNameOfObj, guid)
                if not ok_layer then
                    layer = ""
                    log_warning("Ошибка получения layer для guid: " .. tostring(guid))
                end

                local areagross, areanet = "", ""
                local perimeter, perimeter_openings = "", ""

                -- === Вычисляем размеры ===
                local dx, dy = "", ""
                local ok_poly, polydata = pcall(af_request, "getpoly", {holes=0, givelist=1}, guid)
                if ok_poly and polydata and polydata.poly and polydata.poly[1] then
                    local minx, maxx = nil, nil
                    local miny, maxy = nil, nil
                    for _, pt in ipairs(polydata.poly[1]) do
                        if not minx or pt.x < minx then minx = pt.x end
                        if not maxx or pt.x > maxx then maxx = pt.x end
                        if not miny or pt.y < miny then miny = pt.y end
                        if not maxy or pt.y > maxy then maxy = pt.y end
                    end
                    dx = minx and maxx and string.format("%.0f", (maxx-minx)*1000) or ""
                    dy = miny and maxy and string.format("%.0f", (maxy-miny)*1000) or ""
                else
                    log_warning("Ошибка получения polydata для guid: " .. tostring(guid))
                end

                -- Толщина панели (Z)
                local dz = ""
                if q and q.geo then
                    dz = q.geo.z2 and q.geo.z1 and string.format("%.0f", (q.geo.z2 - q.geo.z1)*1000) or ""
                end

                -- === Площади ===
                if q and q.quant then
                    areagross = string.format("%.2f", q.quant.areagross or 0)
                    areanet = string.format("%.2f", q.quant.areanet or 0)
                end

                -- === Вычисляем периметры ===
                local ok_out, outline = pcall(af_request, "getpoly", {holes=0, givelist=1}, guid)
                if ok_out and outline and outline.poly and #outline.poly > 0 then
                    local plen = PolyLength(outline.poly[1])
                    perimeter = string.format("%.2f", plen)
                else
                    log_warning("Ошибка получения outline для guid: " .. tostring(guid))
                end

                local ok_holes, holes = pcall(af_request, "getpoly", {holes=1, givelist=1}, guid)
                local sum = 0
                if ok_holes and holes and holes.poly and #holes.poly > 0 then
                    for _, pts in ipairs(holes.poly) do
                        sum = sum + PolyLength(pts)
                    end
                    perimeter_openings = string.format("%.2f", sum)
                end

                -- === Вывод значений в ячейки excel ===
                af.LibxlNumToCell(book, wsInv, nrow, 0, n)
                af.LibxlMbsToCell(book, wsInv, nrow, 1, id)
                af.LibxlMbsToCell(book, wsInv, nrow, 2, layer)
                af.LibxlMbsToCell(book, wsInv, nrow, 3, dx)
                af.LibxlMbsToCell(book, wsInv, nrow, 4, dy)
                af.LibxlMbsToCell(book, wsInv, nrow, 5, dz)
                af.LibxlMbsToCell(book, wsInv, nrow, 6, areagross)
                af.LibxlMbsToCell(book, wsInv, nrow, 7, areanet)
                af.LibxlMbsToCell(book, wsInv, nrow, 8, perimeter)
                af.LibxlMbsToCell(book, wsInv, nrow, 9, perimeter_openings)
                nrow = nrow + 1
                n = n + 1
            end
        end
    end


-- Save and open
    if not book:save(outFile) then
        af.RaiseError("Не удалось сохранить: " .. outFile)
    end
    book:release()
    af.ExcelOpen(outFile)
end
--------------------------------------------------------------------------------
-- Register export type
--------------------------------------------------------------------------------
gtblListings[7] = {
    strName           = "Control Card",
    strOnInitFunc     = "OnInitControlCard",
    strOnSaveListFunc = "OnSaveListControlCard",
    nSorting          = 2,
    nAllowSameId      = 1,
    nCollectType      = 5,
}