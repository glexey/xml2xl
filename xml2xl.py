import re
import json
import copy
import glob
import string
import xlbuf
import argparse
import xlsxwriter
import formatters
import xml.etree.cElementTree as ET

KEYWORDS = ('name', 'format', 'ignore', 'comment', 'row', 'col', 'xml',
'xml_select', 'entries', 'prefix', 'suffix', 'separator', 'active',
'autofilter', 'column_formats', 'text', 'sfmt', 'tfmt', 'eval', 'link_to',
'link_to_selector', 'link_id', 'draw_border', 'zoom', 'span', 'xml_nodes',
'json_path', 'visibility_filter', 'cfg', 'xml_select_sheet', 'no_commit',
'xml_filter_sheet')

# characters that aren't allowed in named ranges have to be replaced
TR = string.maketrans(ur"-/[] ", ur"_____")

def arr2str(arr):
    # recursively dump hier list contents to a string, or, when format dictionaries
    # are present instead of strings, provide the array for rich string formatting
    result_arr = []
    result_str = ""
    for x in arr:
        if x is None:
            pass
        elif isinstance(x, basestring):
            if result_str is not None: result_str += x
            result_arr.append(x)
        elif isinstance(x, dict):
            result_str = None
            result_arr.append(x)
        else: # list
            _str, _arr = arr2str(x)
            if result_str is not None:
                if _str is None:
                    result_str = None
                else:
                    result_str += _str
            result_arr += _arr
    return result_str, result_arr

def stub_msg_callback(s):
    return

class Cursor:
    
    # 2D pointer with max coord tracking
    def __init__(self, row=0, col=0):
        self.row = row
        self.col = col
        self.max_row = row
        self.max_col = col

    def update_max(self):
        self.max_row = max(self.row, self.max_row)
        self.max_col = max(self.col, self.max_col)

    def copy(self):
        c = Cursor(self.row, self.col)
        c.max_row = self.max_row
        c.max_col = self.max_col
        return c

class XLName:
    def __init__(self, names = []):
        self.name_map = {}
        for name in names:
            orig_name = name
            while len(name) > 31:
                name = re.sub(r'.*?/|.', '', name, count=1)
            name = re.sub(r'[^\w]+', '.', name)
            if name in self.name_map.values():
                # Failsafe to make sure pages are unique
                import hashlib
                h = hashlib.md5(orig_name).hexdigest()
                name = name[:25] + h[-6:]
            self.name_map[orig_name] = name
    def __getitem__(self, name):
        if name in self.name_map:
            return self.name_map[name]
        else:
            return name

class Sheet:

    # global buffer for cross-sheet data
    cellref = {}

    def __init__(self, cfg, dad, xml):
        self.cfg = cfg
        self.dad = dad
        self.xml = xml
        self.need_url = [] # cells that need urls on this sheet
        self.cell_fmt = {} # keeps track of cell formatting for xlbuf
        self.xlsheet = dad.workbook.add_worksheet(self.dad.xlname[self.cfg['name']])
        self.cellbuf = xlbuf.CellBuffer(dad.workbook, dad.default_fmt)
        self.column_formats = {}
        self.column_widths = {}
        self.column_headers = []
        self.last_visibility_row = 2
        self.filter_column = None
        self.cursor = Cursor(0, 0) # pointer to the next cell to be written

    def write_all(self):
        self.cellbuf.write_all(self.xlsheet, self.cell_fmt)

    def post_process(self):
        for (cell_from, sheet_to_name, link_id) in self.need_url:
            cell_to = Sheet.cellref[(sheet_to_name, link_id)]
            xlref = xlsxwriter.utility.xl_rowcol_to_cell(cell_to.y, cell_to.x, False, False) 
            cell_from.url = "internal:'%s'!%s"%(self.dad.xlname[sheet_to_name], xlref)

    def process(self):

        # Apply column width and default column cell format
        for cfmt in self.cfg.get('column_formats', []):
            if "column_widths" in cfmt:
                # Deprecated, use auto-incremented list instead
                for icol,f in enumerate(cfmt['column_widths']):
                    self.column_widths[icol] = f
            else:
                if "column" in cfmt:
                    # Deprecated, use auto-incremented list instead
                    if "header" in cfmt:
                        raise Exception('Syntax error: both "column" and "header" key present in column_formats list')
                    icol = cfmt['column']
                else:
                    icol = len(self.column_headers)
                    self.column_headers.append({'text': cfmt.get('header', '')})
                    self.column_headers[-1].update(cfmt)
                fmt = self.dad.try_fmt(cfmt, "cell_format")
                if fmt is not None:
                    # Excel will not use column default format as base when cell has *any* other formatting
                    # so we store the column format, to be able to apply cell format on top ourselves
                    self.column_formats[icol] = fmt
                width = cfmt.get('width', None)
                huspath = cfmt.get('hide_unless_select', None)
                if huspath is not None:
                    if self.xml.find(huspath) is None:
                        width = 0
                if width is not None:
                    self.column_widths[icol] = width
        for icol in set(self.column_formats.keys() + self.column_widths.keys()):
            self.xlsheet.set_column(icol, icol, self.column_widths.get(icol, None),
                    self.cellbuf.get_xl_fmt(self.column_formats.get(icol, None)))

        root_entry = self.cfg
        root_entry['xml_nodes'] = [self.xml] # root node is our starting hierarchy
        root_entry['json_path'] = "sheet(%s)"%self.cfg['name'] # keep track of path for debugging
        self.dad.process_entry(self, root_entry, self.cursor, self.cellbuf, self.dad.default_fmt)

        # Make this sheet default active one, if requested
        if self.cfg.get('active', False):
            self.xlsheet.activate()

        # Set zoom level
        self.xlsheet.set_zoom(self.cfg.get('zoom', 100))

        # Add the xl auto filter on the whole range of generated cells, when requested
        if self.cfg.get('autofilter', False):
            self.xlsheet.autofilter(0, 0, self.cursor.max_row, self.cursor.max_col)
            self.xlsheet.freeze_panes(1, 0)

        if self.filter_column is not None:
            self.xlsheet.autofilter(0, self.filter_column, self.cursor.max_row, self.filter_column)
            self.xlsheet.freeze_panes(1, 0)


class XML2XL:

    def __init__(self):
        pass


    def filtercfg_skip(self, entry, filtercfg):
        # return true if we need to skip this entry
        if not isinstance(entry, dict): return False
        if 'ignore' in entry: return True
        if filtercfg is None: return False
        if 'cfg' not in entry: return False
        for cfg in entry['cfg'].split(','):
            if cfg.startswith('!'):
                if filtercfg == cfg[1:]: return True # skip
                return False # negative didn't match: go ahead and process
            else:
                if filtercfg == cfg: break # positive match: go ahead and process
        return True # if nothing matched, skip


    def copy_with_filter(self, a, filtercfg, r=0):
        if isinstance(a, list):
            out = [ self.copy_with_filter(x, filtercfg, r+1) for x in a if not self.filtercfg_skip(x, filtercfg) ]
        elif isinstance(a, dict):
            out = a.copy() # shallow copy first
            for k in a.iterkeys():
                out[k] = self.copy_with_filter(a[k], filtercfg, r+1)
        else: # assume it's either list, or dict, or "literal" (string/int/float)
            out = a
        return out


    def et2xl(self, element_tree, cfg_filename, output_filename, properties = None,
            text_formatter = formatters.xml_strip_formatter, msg_callback = stub_msg_callback,
            filtercfg = None):
        """
        Process XML element tree and write formatted Excel output
        """

        if cfg_filename.endswith('.json'):
            cfg = json.load(open(cfg_filename))
        else:
            config = {}
            execfile(cfg_filename, config)
            try:
                cfg = config['xlmap']
            except:
                raise Exception("'xlmap' not defined in %s"%cfg_filename)
        self.xml = element_tree
        self.text_formatter = text_formatter

        self.cfg = self.copy_with_filter(cfg, filtercfg)

        self.workbook = xlsxwriter.Workbook(output_filename)
        if properties is not None and properties != "":
            if isinstance(properties, basestring):
                a = properties.split(';')
                properties = dict(x.split(':') for x in a)
            self.workbook.set_properties(properties)

        self.default_fmt = self.cfg['formats']['DEFAULT']

        sheets = []
        for cfg in self.cfg['sheets']:
            xss = cfg.get('xml_select_sheet', None)
            # xml_select_sheet allows to select parts of the hierarchy
            # and process each as a separate sheet
            xfs = cfg.get('xml_filter_sheet', None)
            # xml_filter_sheet allows to split single hierarchy into sheets
            # according to the value of a given selector 
            self.xlname = XLName() # identity map
            if xss is not None:
                for mxml in self.xml.findall(xss['select_path']):
                    mcfg = copy.deepcopy(cfg)
                    mcfg['name'] = mxml.findtext(xss['select_name'])
                    sheets.append(Sheet(mcfg, self, mxml))
            elif xfs is not None:
                sheet_names = sorted(set([x.text for x in self.xml.findall(xfs)]))
                self.xlname = XLName(sheet_names)
                for sheet_name in sheet_names:
                    mcfg = copy.deepcopy(cfg)
                    mcfg['name'] = sheet_name
                    sheets.append(Sheet(mcfg, self, self.xml))
            else:
                sheets.append(Sheet(cfg, self, self.xml))

        # First pass, populate all the data
        for sheet in sheets:
            msg_callback("processing sheet '%s'"%sheet.cfg['name'])
            sheet.process()
        
        # Populate links
        msg_callback("populating links")
        for sheet in sheets: sheet.post_process()

        # Dump the buffers
        msg_callback("writing output: " + output_filename)
        for sheet in sheets: sheet.write_all()

        msg_callback("Closing the workbook")
        self.workbook.close()


    def process_entry(self, sheet, entry, cursor, buf, fmt, link_to=None, link_id=None):
        # main [recursive] function that does all the processing

        if not isinstance(entry, dict): raise Exception("Entry is not a dict, instead it is: "%entry)

        # Process cell formatting
        new_fmt = self.try_fmt(entry, 'format')
        if new_fmt is not None:
            fmt = fmt.copy()
            fmt.update(new_fmt)

        if "text" in entry:
            # Simplest case is when we already have the content as a string or array
            # just return it (put it into the cell first if we have the cursor)
            if cursor is not None:
                cell_fmt = sheet.column_formats.get(cursor.col, {}).copy()
                cell_fmt.update(fmt)
                if not entry.get('no_commit', False):
                    cursor.update_max() # Update max location once we've written into the cell
                text, text_arr = arr2str(entry["text"]) # Recursively unpack the array
                if link_to is None:
                    if text is None:
                        buf.cell(cursor.row, cursor.col, text_arr, cell_fmt)
                    else:
                        buf.cell(cursor.row, cursor.col, text, cell_fmt)
                else:
                    if link_id is None: link_id = text
                    #ref = xlsxwriter.utility.xl_rowcol_to_cell(cursor.row, cursor.col, True, True)
                    #self.workbook.define_name(str("%s__%s"%(self.sheet_cfg['name'], link_id)).translate(TR), "='%s'!%s"%(self.sheet_cfg['name'], ref))
                    #self.last_link = str("%s__%s"%(link_to, link_id)).translate(TR)
                    cell = buf.cell(cursor.row, cursor.col, text, cell_fmt)
                    Sheet.cellref[(sheet.cfg['name'], link_id)] = cell # keep the reference to the cell so we can link on 2nd pass
                    sheet.need_url.append((cell, link_to, link_id))
                # merge cells if requested
                span = entry.get('span', None)
                if span is not None:
                    sheet.xlsheet.merge_range(cursor.row, cursor.col, cursor.row+span[0]-1, cursor.col+span[1]-1, '')
            return entry["text"]
        
        try:
            xml_nodes = entry['xml_nodes']
            json_path = entry['json_path']
        except KeyError:
            # Should never happen
            raise Exception("FATAL :: entry xml_node or json_path not defined")

        # Check keywords/syntax for dictionary entries
        for key in entry.keys():
            if key not in KEYWORDS:
                raise Exception('Unrecognized keyword:', key)

        is_cell_text_creator = not ('row' in entry or 'col' in entry)
        child_cursor = None if is_cell_text_creator else cursor
        entry_start_cursor = None if cursor is None else cursor.copy()
        border = entry.get('draw_border', None)

        values = []
        if 'xml' in entry:
            path = entry['xml']
            node_idx = -1 # Last node by default
            imax = None
            make_set = False
            if path.startswith('!'): # Take first found element only
                path = path[1:]
                imax = 1
            if path.startswith('@'): # List unique elements.. default separator of '|'
                path = path[1:]
                make_set = True
                if 'separator' not in entry: entry['separator'] = '|'
            while path.startswith('../'): # Access to the parent's fields
                node_idx -= 1
                path = path[3:]
            apath = path.split("#") # Access to attributes, as opposed to text
            for x in xml_nodes[node_idx].findall(apath[0])[0:imax]:
                if len(apath) == 1:
                    if x.text is None: continue
                    # Format the text with either user-provided or default routine
                    if 'sfmt' in entry:
                        text = formatters.by_name[entry['sfmt']](x.text)
                    else:
                        text = self.text_formatter(x.text)
                else:
                    text = x.get(apath[1])
                if (not make_set) or (text not in values):
                    values.append(text)
        if 'text' in entry:
            values = [entry['text']]
        if 'xml_select' in entry:
            raise Exception("At %s: xml_select entry can only appear directly under the 'entries' list"%json_path)
        if 'entries' in entry:
            # Populate the list of children to iterate over
            child_entries = []
            
            if entry['entries'] == '#column_headers':
                # Special format for column headers population from the list of column formats 
                entry['entries'] = sheet.column_headers
                for e in entry['entries']:
                    if e.get('width', None) == 0 or e['text'] == '':
                        e['no_commit'] = True
            
            for i,child_entry in enumerate(entry['entries']):
                if isinstance(child_entry, dict) and ("xml_select" in child_entry):
                    # expand the xml_select into 0..+inf child entries
                    path = child_entry['xml_select'].replace('%SHEETNAME%', sheet.cfg['name'])
                    if not isinstance(path, basestring):
                        raise Exception("At %s/entries[%d]: xml_select value has to be a string path, instead got %s"%(json_path, i, path))
                    imax = None
                    if path.startswith('!'): # Take first found element only
                        path = path[1:]
                        imax = 1
                    child_xml_nodes = xml_nodes[-1].findall(path)[0:imax]
                    for child_xml_node in child_xml_nodes:
                        new_entry = child_entry.copy()
                        del new_entry['xml_select']
                        new_entry['xml_nodes'] = xml_nodes + [child_xml_node]
                        new_entry['json_path'] = json_path + "/xml_select(%s)"%path
                        child_entries.append(new_entry)
                else:
                    # shortcuts for the user
                    if isinstance(child_entry, basestring) or isinstance(child_entry, list):
                        child_entry = {"text": child_entry}
                    # append list of child entries to process, at same XML hierarchy
                    child_entry['xml_nodes'] = xml_nodes
                    child_entry['json_path'] = json_path + "/entries[%d]"%i
                    child_entries.append(child_entry)

            # Now walk over children
            for i,child_entry in enumerate(child_entries):
                values.append(self.process_entry(sheet, child_entry, child_cursor, buf, fmt))
                if i == len(child_entries) - 1:
                    if border is not None:
                        lt = [entry_start_cursor.row, entry_start_cursor.col]
                        rb = [cursor.max_row, cursor.max_col]
                        buf.draw_range_border(lt, rb, border.get("type", None), border.get("color", None))
                    if False and 'visibility_filter' in entry: # Disabled for now, need to fix last_link to work
                        f = entry['visibility_filter']
                        formula = "=SUBTOTAL(103, %s)"%self.last_link
                        formula2 = "=%s"%xlsxwriter.utility.xl_rowcol_to_cell(self.last_visibility_row, f['column'], True, True)
                        for r in range(self.last_visibility_row, cursor.row + 1):
                            buf.cell(r, f['column'], formula)
                            formula = formula2
                        self.last_visibility_row = cursor.row + 1
                        self.filter_column = f['column']
                else:
                    self.move_cursor(entry, cursor)

        if is_cell_text_creator:
            str_eval = entry.get('eval', 'x') # Arbitrary user transformation with 'x' as an input variable
            tfmt = self.try_fmt(entry, 'tfmt') # Rich string formatting
            prefix = entry.get('prefix', None)
            suffix = entry.get('suffix', None)
            separator = entry.get('separator', None)
            result_str = []
            for i,x in enumerate(values):
                if prefix is not None:
                    if tfmt is not None: result_str.append(tfmt)
                    result_str.append(prefix)
                try:
                    v = eval(str_eval)
                except:
                    print "str_eval=%s x=%s"%(str_eval, x)
                    raise
                if v is not None and v != "":
                    if tfmt is not None: result_str.append(tfmt)
                    result_str.append(v)
                if suffix is not None:
                    if tfmt is not None: result_str.append(tfmt)
                    result_str.append(suffix)
                if separator is not None and i < len(values) - 1:
                    if tfmt is not None: result_str.append(tfmt)
                    result_str.append(separator)
            leaf_entry = entry.copy()
            leaf_entry["text"] = result_str
            leaf_entry["json_path"] += "/text"
            link_id = entry.get('link_id', None)
            if link_id is not None: link_id = xml_nodes[-1].findtext(link_id)
            if 'link_to_selector' in entry:
                entry['link_to'] = xml_nodes[-1].findtext(entry['link_to_selector'])
            return self.process_entry(sheet, leaf_entry, cursor, buf, fmt, entry.get('link_to', None), link_id)


    def try_fmt(self, dic, key):
        fmt = dic.get(key, None)
        if fmt is not None:
            if isinstance(fmt, basestring): fmt = self.cfg['formats'][fmt]
        return fmt


    def move_cursor(self, entry, cursor):
        for key in ("row", "col"):
            if key in entry:
                if cursor is None: raise Exception("No cursor at " + entry['json_path'])
                a = entry[key]
                if a.startswith('+'): setattr(cursor, key, getattr(cursor, key) + int(a[1:]))
                elif a.startswith('-'): setattr(cursor, key, getattr(cursor, key) - int(a[1:]))
                else: setattr(cursor, key, int(a))


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("-x", "--xml", required=True, nargs='*', metavar='FILENAME', help="Input .xml file name(s)")
    parser.add_argument("-c", "--cfg", required=True, help="Config (.json) file name")
    parser.add_argument("-o", "--output", help="Output (.xlsx) file name")
    parser.add_argument("-C", "--filtercfg", help="Field filter config")
    parser.add_argument("-p", "--properties", help="Set of properties to embed into the doc, in the form: prop1:blah blah;prop2:meh meh")
    args = parser.parse_args()

    filenames = []
    for fnglob in args.xml:
        filenames += glob.glob(fnglob)

    if args.output is None:
        args.output = re.sub(r'(?:\.\w+)?$', '.xlsx', filenames[0])
    else:
        if not args.output.endswith(".xlsx"):
            raise Exception("Output file name should have .xlsx extension")

    top = ET.Element('top')
    for fxml in filenames:
        element_tree = ET.parse(fxml)
        for elem in element_tree.getroot():
            # Merge all given XMLs under under the same root
            top.append(elem)

    print "Writing:", args.output 
    XML2XL().et2xl(top, args.cfg, args.output, properties = args.properties, filtercfg = args.filtercfg)
