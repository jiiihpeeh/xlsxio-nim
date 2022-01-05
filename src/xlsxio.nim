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
    #VersionError* = object of ValueError
    XlsxIOSkip* = enum
        None = 0x00,
        EmptyRows = 0x01,
        EmptyCells = 0x02,
        AllEmpty = 0x03,
        ExtraCells = 0x04,
        HiddenRows = 0x08
    XlsxAddMode* = enum
        Int, DateTime
    XlsxCellMode* {.pure.} = enum
        String, Integer, Float
    XlsxCell* = tuple[state: bool, value: string]
#[     ReadProcessCellCallback*[T] = proc (row: int; col: int; value: string; callbackdata: var T)
    ReadProcessRowCallback*[T] = proc (row: int; maxcol: int; callbackdata: var T) ]#


{.push warnings: off.}
proc skippables(): seq[int]{.compiletime.} =
    for i in XlsxIOSkip:
        result.add i.int
{.pop.}

const
    skipNums = skippables()
    epochOffsetf = -2209075200'f64
    epochOffseti = -2209075200'i64

proc readGetVersion*(): XlsxioVersion =
    var pmajor, pminor, pmicro: cint
    xlsxio_read_core.xlsxioread_get_version(pmajor.addr, pminor.addr, pmicro.addr)
    return XlsxioVersion(major: pmajor.int, minor: pminor.int,
            micro: pmicro.int)

proc `$`*(version: XlsxioVersion): string =
    ##Formats version info into a string
    return $ version.major & "." & $ version.minor & "." & $ version.micro

#[ This is  useless
    proc readGetVersionString*(): string =
    let v = xlsxio_read_core.xlsxioread_get_version_string()
    if not v.isNil:
        return $ v
    else:
        raise newException(VersionError, "Can not get version") ]#


proc readOpen*(filename: string): Xlsxioreader =
    ##Opens a spreadsheet for reading. Returns a handle.
    var reader = xlsxio_read_core.xlsxioread_open(filename.cstring)
    #defer xlsxio_read_core.xlsxioread_close(handle)
    if not reader.isNil:
        return reader
    else:
        raise newException(IOError, "Can not open a file")

proc readOpenFilehandle*(filehandle: File): Xlsxioreader =
    ##Opens a spreadsheet file handle for reading. Returns a handle.
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

proc readOpenMemory*(data: var string; datalen: int;
        freedata: bool = true): Xlsxioreader =
    var reader = xlsxio_read_core.xlsxioread_open_memory(data.addr,
            datalen.uint64, boolToCint(freedata))
    if not reader.isNil:
        return reader
    else:
        raise newException(IOError, "Can not handle data")

proc readClose*(handle: Xlsxioreader) =
    ##Closes a handle.
    xlsxio_read_core.xlsxioread_close(handle)



proc readSheetlistOpen*(handle: Xlsxioreader): Xlsxioreadersheetlist =
    var sheethandle = xlsxio_read_core.xlsxioread_sheetlist_open(handle)
    if sheethandle.isNil:
        raise newException(IOError, "Can not read sheets (handle: nil)")
    else:
        return sheethandle

iterator readSheets*(handle: Xlsxioreader): string =
    if handle.isNil:
        raise newException(IOError, "Can not read a file (handle: nil)")
    var listHandle = xlsxioreadSheetlistOpen(handle)
    while true:
        let sheetname = xlsxio_read_core.xlsxioread_sheetlist_next(listHandle)
        if sheetname.isNil:
            xlsxio_read_core.xlsxioread_sheetlist_close(listHandle)
            break
        else:
            yield $ sheetname

iterator readSheets*(handle: Xlsxioreadersheetlist): string =
    if handle.isNil:
        raise newException(IOError, "Can not read sheets (handle: nil)")
    while true:
        let sheetname = xlsxio_read_core.xlsxioread_sheetlist_next(handle)
        if sheetname.isNil:
            xlsxio_read_core.xlsxioread_sheetlist_close(handle)
            break
        else:
            yield $ sheetname

proc sheets*(handle: Xlsxioreader): seq[string] =
    ## Lists read handle's sheets.
    return toSeq(readSheets(handle))


proc hasSheet*(handle: Xlsxioreader; name: string): bool =
    # Checks if a read handle has a sheet
    var listhandle = xlsxioreadSheetlistOpen(handle)
    #defer: xlsxio_read_core.xlsxioread_sheetlist_close(listhandle)
    for s in readSheets(listhandle):
        if name == s:
            xlsxio_read_core.xlsxioread_sheetlist_close(listhandle)
            return true
    xlsxio_read_core.xlsxioread_sheetlist_close(listhandle)
    return false

proc len*(handle: Xlsxioreader): int =
    var listhandle = xlsxioreadSheetlistOpen(handle)
    var count = 0
    #defer: xlsxio_read_core.xlsxioread_sheetlist_close(listhandle)
    for s in readSheets(listhandle):
        count += 1
    return count

proc readSheetOpen*(handle: Xlsxioreader; sheetname: string;
        skip: XlsxIOSkip = None): Xlsxioreadersheet =
    var reader = xlsxio_read_core.xlsxioread_sheet_open(handle,
            sheetname.cstring, skip.cuint)
    if not reader.isNil:
        return reader
    else:
        raise newException(IOError, "Can not read sheets (handle: nil)")

proc readSheetOpen*(handle: Xlsxioreader; sheetindex: int;
        skip: XlsxIOSkip = None): Xlsxioreadersheet =
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



proc readSheetLastRowIndex*(sheethandle: Xlsxioreadersheet): int =
    let lastIndex = xlsxio_read_core.xlsxioread_sheet_last_row_index(sheethandle)
    return lastIndex.int


proc readSheetLastColumnIndex*(sheethandle: Xlsxioreadersheet): int =
    let lastIndex = xlsxio_read_core.xlsxioread_sheet_last_column_index(sheethandle)
    return lastIndex.int


proc readSheetNextRow*(sheethandle: Xlsxioreadersheet): int =
    let b = xlsxio_read_core.xlsxioread_sheet_next_row(sheethandle)
    return b.int

proc readSheetNextCell*(sheethandle: Xlsxioreadersheet): XlsxCell =
    var result : XlsxCell 
    var cell = xlsxio_read_core.xlsxioread_sheet_next_cell(sheethandle)
    if not cell.isNil:
        result = (true, $cell)
    xlsxio_read_core.xlsxioread_free(cell)

    return result


var cellstringNext : cstring 

proc readSheetNextCellString*(sheethandle: Xlsxioreadersheet;
        cellstring: var string): int =
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_string(sheethandle,
            cellstringNext.addr)
    cellstring = $ cellstringNext
    xlsxio_read_core.xlsxioread_free(cellstringNext)
    return status.int

proc readSheetNextCellInt*(sheethandle: Xlsxioreadersheet;
        cellint: var int64): int =
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_int(sheethandle, cellint.addr)
    return status.int


proc readSheetNextCellFloat*(sheethandle: Xlsxioreadersheet;
        cellfloat: var float64): int =
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_float(sheethandle,
            cellfloat.addr)
    return status.int

proc readSheetNextCellEpoch*(sheethandle: Xlsxioreadersheet;
        cellint: var int64): int =
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_datetime(
            sheethandle, cellint.addr)
    return status.int

proc readSheetNextCellTime*(sheethandle: Xlsxioreadersheet;
        celltime: var Time): int =
    var cell: int64
    let status = xlsxio_read_core.xlsxioread_sheet_next_cell_datetime(
            sheethandle, cell.addr)
    celltime = fromUnix(cell)
    return status.int


{.push warnings: off.}
proc readSheetFlags*(sheethandle: Xlsxioreadersheet): XlsxIOSkip =
    let f = xlsxio_read_core.xlsxioread_sheet_flags(sheethandle).int
    if skipNums.contains f:
        return f.XlsxIOSkip
    else:
        return None
{.pop.}


template modeParse(argument: untyped, mode: XlsxCellMode): untyped =
    case mode
    of String:
        argument
    else:
        argument


iterator readSheetRows*(sheethandle: Xlsxioreadersheet): seq[string] =
    discard readSheetNextCell(sheethandle)
    var row = newSeq[string](0)
    while true:
        while true:
            let s = readSheetNextCell(sheethandle)
            if s.state:
                row.add s.value
            else:
                break
        yield row
        if readSheetNextRow(sheethandle) == 0:
            break
        row = newSeq[string](0)

iterator readSheetRowsInt*(sheethandle: Xlsxioreadersheet): seq[int] =
    discard readSheetNextCell(sheethandle)
    var row = newSeq[int](0)
    while true:
        while true:
            let s = readSheetNextCell(sheethandle)
            if s.state:
                try:
                    row.add s.value.parseInt
                except:
                    discard
            else:
                break
        yield row
        if readSheetNextRow(sheethandle) == 0:
            break
        row = newSeq[int](0)

iterator readSheetRowsFloat*(sheethandle: Xlsxioreadersheet): seq[float] =
    discard readSheetNextCell(sheethandle)
    var row = newSeq[float](0)
    while true:
        while true:
            let s = readSheetNextCell(sheethandle)
            if s.state:
                try:
                    row.add s.value.parseFloat
                except:
                    discard
            else:
                break
        yield row
        if readSheetNextRow(sheethandle) == 0:
            break
        row = newSeq[float](0)

iterator readSheetRowsEpoch*(sheethandle: Xlsxioreadersheet): seq[float] =
    discard readSheetNextCell(sheethandle)
    var row = newSeq[float](0)
    while true:
        while true:
            let s = readSheetNextCell(sheethandle)
            if s.state:
                try:
                    row.add s.value.parseFloat + epochOffsetf
                except:
                    discard
            else:
                break
        yield row
        if readSheetNextRow(sheethandle) == 0:
            break
        row = newSeq[float](0)

iterator readSheetRowsTime*(sheethandle: Xlsxioreadersheet): seq[Time] =
    discard readSheetNextCell(sheethandle)
    var row = newSeq[Time](0)
    while true:
        while true:
            let s = readSheetNextCell(sheethandle)
            if s.state:
                try:
                    row.add fromUnixFloat(s.value.parseFloat + epochOffsetf)
                except:
                    discard
            else:
                break
        yield row
        if readSheetNextRow(sheethandle) == 0:
            break
        row = newSeq[Time](0)

proc readSheetIntoArray*(sheethandle: Xlsxioreadersheet): seq[seq[string]] =
    result = newSeq[seq[string]](0)
    for r in readSheetRows(sheethandle):
        result.add r



proc readSheetIntoTable(sheethandle: Xlsxioreadersheet) =
    discard

#[ proc readSheetIntoOrderedTable(sheethandle: Xlsxioreadersheet)=
    discard readSheetNextCell(sheet)
    var sd = initOrderedTable[int,seq[string]](0)
    var row = 1
    var sek = newSeq[string](0)
    while true:
        let s = readSheetNextCell(sheet)
        if s.state:
            sek.add s.value
        else:
            sd[sd.len + 1] = sek
            sek = newSeq[string](0)
            if readSheetNextRow(sheet) == 0:
                break
            ]#
proc writeGetVersion*(): XlsxioVersion =
    var pmajor, pminor, pmicro: cint
    xlsxio_write_core.xlsxiowrite_get_version(pmajor.addr, pminor.addr, pmicro.addr)
    return XlsxioVersion(major: pmajor.int, minor: pminor.int,
            micro: pmicro.int)


#[
    This is  useless
     proc writeGetVersionString*(): string =
    let v = xlsxio_write_core.xlsxiowrite_get_version_string()
    if not v.isNil:
        return $ v
    else:
        raise newException(VersionError, "Can not get a version")

 ]#
proc writeOpen*(filename: string; sheetname: string): Xlsxiowriter =
    var writeHandle = xlsxio_write_core.xlsxiowrite_open(filename.cstring,
            sheetname.cstring)
    if not writeHandle.isNil:
        return writeHandle
    else:
        raise newException(IOError, "Can not get a file handle")



proc writeClose*(handle: Xlsxiowriter) =
    let c = xlsxio_write_core.xlsxiowrite_close(handle)
    if c != 0.cint:
        raise newException(IOError, "Can not close a file handle")


proc writeSetDetectionRows*(handle: Xlsxiowriter; rows: int) =
    xlsxio_write_core.xlsxiowrite_set_detection_rows(handle, rows.csize_t)

proc writeSetRowHeight*(handle: Xlsxiowriter; height: int) =
    xlsxio_write_core.xlsxiowrite_set_row_height(handle, height.csize_t)

proc writeAddColumn*(handle: Xlsxiowriter; name: string; width: int) =
    xlsxio_write_core.xlsxiowrite_add_column(handle, name.cstring, width.cint)

proc writeAddCellString*(handle: Xlsxiowriter; value: string) =
    xlsxio_write_core.xlsxiowrite_add_cell_string(handle, value.cstring)

proc writeAddCellInt*(handle: Xlsxiowriter; value: int64) =
    xlsxio_write_core.xlsxiowrite_add_cell_int(handle, value)

proc writeAddCellFloat*(handle: Xlsxiowriter; value: float64) =
    xlsxio_write_core.xlsxiowrite_add_cell_float(handle, value)

proc writeAddCellEpoch*(handle: Xlsxiowriter; value: int64) =
    xlsxio_write_core.xlsxiowrite_add_cell_datetime(handle, value)


proc writeAddCellTime*(handle: Xlsxiowriter; value: Time) =
    let epoch = toUnix(value)
    xlsxio_write_core.xlsxiowrite_add_cell_datetime(handle, epoch)

proc writeNextRow*(handle: Xlsxiowriter) =
    xlsxio_write_core.xlsxiowrite_next_row(handle)

proc writeAddCell*(handle: Xlsxiowriter; value: string) =
    xlsxio_write_core.xlsxiowrite_add_cell_string(handle, value.cstring)


proc writeAddCell*(handle: Xlsxiowriter; value: int64,
        mode: XlsxAddMode = Int) =
    case mode
    of Int:
        xlsxio_write_core.xlsxiowrite_add_cell_int(handle, value)
    of DateTime:
        xlsxio_write_core.xlsxiowrite_add_cell_datetime(handle, value)


proc writeAddCell*(handle: Xlsxiowriter; value: float64) =
    xlsxio_write_core.xlsxiowrite_add_cell_float(handle, value)
