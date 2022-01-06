##  This is a bit higher level wrapper for xlsxio. It automates memory freeing,
##  converts Nil errors into exceptions, provides few iterators/ procedures
##  and turns some constants into enums. Lower level access is available in src/xlsxio/ directory.
## 
## 
##  In order to keep it more readable this higher level wrapper omits xlsxio prefix in procedure names. Due to
##  quite distinctively named types namespace collisions/confusion should hopefully not happen. 
##  Prefix read/write correspond to their respective xlsxio counterparts. Time conversion omits
##  dateTime naming convention and uses epoch (unix) and Time from the standard library.
## 
##  This wrapper tries to keep all useful procedures available so custom iterators can be easily made.
## 
runnableExamples:
    #should be import xlsxio
    let handle = open("calc.xlsx") 
    let xlsx = readOpenFilehandle(handle) #alternatively let xlsx = readOpen("calc.xlsx")
    let sheet = xlsx.readSheetOpen("Sheet 1") # or by index xlsx.readSheetOpen(1)
    #echoes sheet content row by row in seq[string]
    for i in readSheetRows(sheet):
        echo i
    xlsx.readClose()
    handle.close()
 
import
    xlsxio/xlsxio_read_core,
    xlsxio/xlsxio_write_core,
    std/sequtils,
    std/enumutils,
    std/strutils,
    std/times
{.passL: "-lxlsxio_read".}
{.passL: "-lxlsxio_write".}
type
    XlsxIOVersion* = object
        major, minor, micro: int
    XlsxIOSkip* = enum
        None = 0x00,
        EmptyRows = 0x01,
        EmptyCells = 0x02,
        AllEmpty = 0x03,
        ExtraCells = 0x04,
        HiddenRows = 0x08
    XlsxAddMode* = enum
        Integer, TimeInfo
    XlsxCell* = tuple[state: bool, value: string]



{.push warnings: off.}
proc skippables(): seq[int]{.compiletime.} =
    for i in XlsxIOSkip:
        result.add i.int
{.pop.}

const
    skipNums = skippables()
    epochOffsetF = -2209075200'f64
    #epochOffseti = -2209075200'i64

proc readGetVersion*(): XlsxioVersion =
    ## Returns version info in object format
    var pmajor, pminor, pmicro: cint
    xlsxio_read_core.xlsxioread_get_version(pmajor.addr, pminor.addr, pmicro.addr)
    return XlsxioVersion(major: pmajor.int, minor: pminor.int,
            micro: pmicro.int)

proc `$`*(version: XlsxioVersion): string =
    ## Formats version info into a string
    runnableExamples:
        import xlsxio
        let version = $ readGetVersion
        # Gives something like "0.20.2"
    return $ version.major & "." & $ version.minor & "." & $ version.micro


proc readOpen*(filename: string): XlsxioReader =
    ## Opens a spreadsheet for reading. Returns a handle.
    var reader = xlsxio_read_core.xlsxioread_open(filename.cstring)
    if not reader.isNil:
        return reader
    else:
        raise newException(IOError, "Can not open a file")

proc readOpenFilehandle*(filehandle: File): XlsxioReader =
    ## Opens a spreadsheet file handle for reading. Returns a handle.
    runnableExamples:
        let handle = open("calc.xlsx")
        let xlsx = readOpenFilehandle(handle)
    var osHandle = getOsFileHandle(filehandle)
    var reader = xlsxio_read_core.xlsxioread_open_filehandle(osHandle)
    if not reader.isNil:
        return reader
    else:
        raise newException(IOError, "Can not open a filehandle")

template boolToCint(flag: bool): cint =
    var intflag = 0.cint
    if flag == true:
        intflag = 1
    intflag

proc readOpenMemory*(data: ptr; datalen: int;
        freedata: bool = true): XlsxioReader =
    ## Not tested
    var reader = xlsxio_read_core.xlsxioread_open_memory(data,
            datalen.uint64, boolToCint(freedata))
    if not reader.isNil:
        return reader
    else:
        raise newException(IOError, "Can not handle data")

proc readClose*(handle: XlsxioReader) =
    ## Closes a handle.
    xlsxio_read_core.xlsxioread_close(handle)



proc readSheetlistOpen*(handle: XlsxioReader): XlsxioReaderSheetList =
    var handle = xlsxio_read_core.xlsxioread_sheetlist_open(handle)
    if handle.isNil:
        raise newException(IOError, "Can not read sheets (handle: nil)")
    else:
        return handle

iterator readSheets*(handle: XlsxioReader): string =
    ## Iterates over aviable sheets
    if handle.isNil:
        raise newException(IOError, "Can not read a file (handle: nil)")
    var listHandle = readSheetlistOpen(handle)
    while true:
        let sheetname = xlsxio_read_core.xlsxioread_sheetlist_next(listHandle)
        if sheetname.isNil:
            xlsxio_read_core.xlsxioread_sheetlist_close(listHandle)
            break
        else:
            yield $ sheetname

iterator readSheets*(handle: XlsxioReaderSheetList): string =
    ## Iterates over aviable sheets
    if handle.isNil:
        raise newException(IOError, "Can not read sheets (handle: nil)")
    while true:
        let sheetname = xlsxio_read_core.xlsxioread_sheetlist_next(handle)
        if sheetname.isNil:
            xlsxio_read_core.xlsxioread_sheetlist_close(handle)
            break
        else:
            yield $ sheetname

proc sheets*(handle: XlsxioReader): seq[string] =
    ## Lists read handle's sheets.
    return toSeq(readSheets(handle))


proc hasSheet*(handle: XlsxioReader; name: string): bool =
    ## Checks if a read handle has a sheet
    var listhandle = readSheetlistOpen(handle)
    for s in readSheets(listhandle):
        if name == s:
            xlsxio_read_core.xlsxioread_sheetlist_close(listhandle)
            return true
    xlsxio_read_core.xlsxioread_sheetlist_close(listhandle)
    return false

proc len*(handle: XlsxioReader): int =
    ##Returns sheet count
    var listhandle = readSheetlistOpen(handle)
    var count = 0
    for s in readSheets(listhandle):
        count += 1
    return count

proc readSheetOpen*(handle: XlsxioReader; sheetname: string;
        skip: XlsxIOSkip = None): XlsxioReaderSheet =
    ##Opens a sheet. Takes ignore options.
    var reader = xlsxio_read_core.xlsxioread_sheet_open(handle,
            sheetname.cstring, skip.cuint)
    if not reader.isNil:
        return reader
    else:
        raise newException(IOError, "Can not read sheets (handle: nil)")

proc readSheetOpen*(handle: XlsxioReader; sheetindex: int;
        skip: XlsxIOSkip = None): XlsxioReaderSheet =
    ##Opens a sheet for a given index. Takes ignore options. Indexing is 1 based.
    var index = 1
    var sheetname: string
    var found = false
    for s in readSheets(handle):
        if index == sheetindex:
            sheetname = s
            found = true
            break
        index += 1
    if not found:
        raise newException(ValueError, "Index out of bounds")
    var reader = xlsxio_read_core.xlsxioread_sheet_open(handle,
            sheetname.cstring, skip.cuint)
    if not reader.isNil:
        return reader
    else:
        raise newException(IOError, "Can not read sheets (handle: nil)")



proc readSheetLastRowIndex*(handle: XlsxioReaderSheet): int =
    let lastIndex = xlsxio_read_core.xlsxioread_sheet_last_row_index(handle)
    return lastIndex.int


proc readSheetLastColumnIndex*(handle: XlsxioReaderSheet): int =
    let lastIndex = xlsxio_read_core.xlsxioread_sheet_last_column_index(handle)
    return lastIndex.int


proc readSheetNextRow*(handle: XlsxioReaderSheet): int =
    let b = xlsxio_read_core.xlsxioread_sheet_next_row(handle)
    return b.int

proc readSheetNextCell*(handle: XlsxioReaderSheet): XlsxCell =
    var cell = xlsxio_read_core.xlsxioread_sheet_next_cell(handle)
    if not cell.isNil:
        result = (true, $cell)
    else:
        result = (false, "") 
    xlsxio_read_core.xlsxioread_free(cell)

    return result


var cellstringNext : cstring 

proc readSheetNextCellString*(handle: XlsxioReaderSheet;
        cellstring: var string): int =
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_string(handle,
            cellstringNext.addr)
    cellstring = $ cellstringNext
    xlsxio_read_core.xlsxioread_free(cellstringNext)
    return status.int

proc readSheetNextCellInt*(handle: XlsxioReaderSheet;
        cellint: var int64): int =
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_int(handle, cellint.addr)
    return status.int


proc readSheetNextCellFloat*(handle: XlsxioReaderSheet;
        cellfloat: var float64): int =
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_float(handle,
            cellfloat.addr)
    return status.int

proc readSheetNextCellEpoch*(handle: XlsxioReaderSheet;
        cellint: var int64): int =
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_datetime(
            handle, cellint.addr)
    return status.int

proc readSheetNextCellTime*(handle: XlsxioReaderSheet;
        celltime: var Time): int =
    var cell: int64
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_datetime(
            handle, cell.addr)
    celltime = fromUnix(cell)
    return status.int


{.push warnings: off.}
proc readSheetFlags*(handle: XlsxioReaderSheet): XlsxIOSkip =
    let f = xlsxio_read_core.xlsxioread_sheet_flags(handle).int
    if skipNums.contains f:
        return f.XlsxIOSkip
    else:
        return None
{.pop.}


iterator readSheetRows*(handle: XlsxioReaderSheet): seq[string] =
    discard readSheetNextCell(handle)
    var row = newSeq[string](0)
    while true:
        while true:
            let s = readSheetNextCell(handle)
            if s.state:
                row.add s.value
            else:
                break
        yield row
        if readSheetNextRow(handle) == 0:
            break
        row = newSeq[string](0)

iterator readSheetRowsInt*(handle: XlsxioReaderSheet): seq[int] =
    discard readSheetNextCell(handle)
    var row = newSeq[int](0)
    while true:
        while true:
            let s = readSheetNextCell(handle)
            if s.state:
                try:
                    row.add s.value.parseInt
                except:
                    discard
            else:
                break
        yield row
        if readSheetNextRow(handle) == 0:
            break
        row = newSeq[int](0)

iterator readSheetRowsFloat*(handle: XlsxioReaderSheet): seq[float] =
    discard readSheetNextCell(handle)
    var row = newSeq[float](0)
    while true:
        while true:
            let s = readSheetNextCell(handle)
            if s.state:
                try:
                    row.add s.value.parseFloat
                except:
                    discard
            else:
                break
        yield row
        if readSheetNextRow(handle) == 0:
            break
        row = newSeq[float](0)

iterator readSheetRowsEpoch*(handle: XlsxioReaderSheet): seq[float] =
    discard readSheetNextCell(handle)
    var row = newSeq[float](0)
    while true:
        while true:
            let s = readSheetNextCell(handle)
            if s.state:
                try:
                    row.add s.value.parseFloat + epochOffsetF
                except:
                    discard
            else:
                break
        yield row
        if readSheetNextRow(handle) == 0:
            break
        row = newSeq[float](0)

iterator readSheetRowsTime*(handle: XlsxioReaderSheet): seq[Time] =
    discard readSheetNextCell(handle)
    var row = newSeq[Time](0)
    while true:
        while true:
            let s = readSheetNextCell(handle)
            if s.state:
                try:
                    row.add fromUnixFloat(s.value.parseFloat + epochOffsetF)
                except:
                    discard
            else:
                break
        yield row
        if readSheetNextRow(handle) == 0:
            break
        row = newSeq[Time](0)

proc readSheetIntoArray*(handle: XlsxioReaderSheet): seq[seq[string]] =
    result = newSeq[seq[string]](0)
    for r in readSheetRows(handle):
        result.add r

           
proc writeGetVersion*(): XlsxioVersion =
    var pmajor, pminor, pmicro: cint
    xlsxio_write_core.xlsxiowrite_get_version(pmajor.addr, pminor.addr, pmicro.addr)
    return XlsxioVersion(major: pmajor.int, minor: pminor.int,
            micro: pmicro.int)


proc writeOpen*(filename: string; sheetname: string): XlsxioWriter =
    var writeHandle = xlsxio_write_core.xlsxiowrite_open(filename.cstring,
            sheetname.cstring)
    if not writeHandle.isNil:
        return writeHandle
    else:
        raise newException(IOError, "Can not get a file handle")



proc writeClose*(handle: XlsxioWriter) =
    let c = xlsxio_write_core.xlsxiowrite_close(handle)
    if c != 0.cint:
        raise newException(IOError, "Can not close a file handle")


proc writeSetDetectionRows*(handle: XlsxioWriter; rows: int) =
    xlsxio_write_core.xlsxiowrite_set_detection_rows(handle, rows.csize_t)

proc writeSetRowHeight*(handle: XlsxioWriter; height: int) =
    xlsxio_write_core.xlsxiowrite_set_row_height(handle, height.csize_t)

proc writeAddColumn*(handle: XlsxioWriter; name: string; width: int) =
    xlsxio_write_core.xlsxiowrite_add_column(handle, name.cstring, width.cint)

proc writeAddCellString*(handle: XlsxioWriter; value: string) =
    xlsxio_write_core.xlsxiowrite_add_cell_string(handle, value.cstring)

proc writeAddCellInt*(handle: XlsxioWriter; value: int64) =
    xlsxio_write_core.xlsxiowrite_add_cell_int(handle, value)

proc writeAddCellFloat*(handle: XlsxioWriter; value: float64) =
    xlsxio_write_core.xlsxiowrite_add_cell_float(handle, value)

proc writeAddCellEpoch*(handle: XlsxioWriter; value: int64) =
    xlsxio_write_core.xlsxiowrite_add_cell_datetime(handle, value)


proc writeAddCellTime*(handle: XlsxioWriter; value: Time) =
    let epoch = toUnix(value)
    xlsxio_write_core.xlsxiowrite_add_cell_datetime(handle, epoch)

proc writeAddCell*(handle: XlsxioWriter; value: Time) =
    let epoch = toUnix(value)
    xlsxio_write_core.xlsxiowrite_add_cell_datetime(handle, epoch)

proc writeNextRow*(handle: XlsxioWriter) =
    xlsxio_write_core.xlsxiowrite_next_row(handle)

proc writeAddCell*(handle: XlsxioWriter; value: string) =
    xlsxio_write_core.xlsxiowrite_add_cell_string(handle, value.cstring)


proc writeAddCell*(handle: XlsxioWriter; value: int64,
        mode: XlsxAddMode = Integer) =
    case mode
    of Integer:
        xlsxio_write_core.xlsxiowrite_add_cell_int(handle, value)
    of TimeInfo:
        xlsxio_write_core.xlsxiowrite_add_cell_datetime(handle, value)


proc writeAddCell*(handle: XlsxioWriter; value: float64) =
    xlsxio_write_core.xlsxiowrite_add_cell_float(handle, value)
