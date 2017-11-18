def isstr(a): return isinstance(a, basestring)
def isdict(a): return isinstance(a, dict)

class OneCell:

    def __init__(self, y, x, val, fmt_dict, comment, ref, url):
        self.x = x
        self.y = y
        self.val = val
        self.ref = ref
        self.url = url
        self.calc_val = val
        self.comment = comment
        if fmt_dict is None: fmt_dict = {}
        self.fmt_dict = fmt_dict.copy()


class CellBuffer:

    def __init__(self, workbook, default_fmt_dict):
        self.buf = {}
        self.fmt_buf = {}
        self.default_fmt_dict = default_fmt_dict 
        self.workbook = workbook
        self.row_format = {}
        self.num_urls = 0


    def cell(self, y, x, val = None, fmt_dict = None, comment = None, ref = None, url = None):
        # create the cell, if needed
        if (y,x) not in self.buf:
            self.buf[y,x] = OneCell(y, x, val, self.default_fmt_dict, comment, ref, url)

        # update value and formats
        if val is not None:
            self.buf[y,x].val = val
            self.buf[y,x].calc_val = val

        if fmt_dict is not None:
            self.buf[y,x].fmt_dict.update(self.expand_borders(fmt_dict))

        return self.buf[y,x]


    def get_xl_fmt(self, fmt_dict):
        if fmt_dict is None or fmt_dict == {}: return None
        def dict2hash(d):
            return tuple([(k, d[k]) for k in sorted(d.keys())])
        key = dict2hash(fmt_dict)
        if key in self.fmt_buf:
            fmt = self.fmt_buf[key]
        else:
            fmt = self.workbook.add_format(fmt_dict)
            self.fmt_buf[key] = fmt
        return fmt


    def expand_borders(self, fmt):
        # convert "whole" border format to 4 separate "side" ones
        # so that those could be overwritten individually by draw_range_border()
        efmt = fmt.copy()
        for i in ('', '_color'):
            if 'border'+i in efmt:
                for side in ('left', 'right', 'top', 'bottom'):
                    efmt[side+i] = efmt['border'+i]
                del efmt['border'+i]
        return efmt


    def draw_range_border(self, corner1, corner2, btype=1, bcolor="black"):
        r1, c1 = corner1
        r2, c2 = corner2
        if (r1 > r2): r1, r2 = r2, r1
        if (c1 > c2): c1, c2 = c2, c1
        for row in (r1, r2):
            side = 'top' if row == r1 else 'bottom'
            fmt1 = {side: btype, side+'_color': bcolor}
            for col in range(c1, c2 + 1):
                fmt2 = fmt1.copy()
                if col == c1: fmt2.update({'left': btype, 'left_color': bcolor})
                if col == c2: fmt2.update({'right': btype, 'right_color': bcolor})
                self.cell(row, col, fmt_dict=fmt2)
        for col in (c1, c2):
            side = 'left' if col == c1 else 'right'
            fmt1 = {side: btype, side+'_color': bcolor}
            for row in range(r1, r2 + 1):
                fmt2 = fmt1.copy()
                if row == r1: fmt2.update({'top': btype, 'top_color': bcolor})
                if row == r2: fmt2.update({'bottom': btype, 'bottom_color': bcolor})
                self.cell(row, col, fmt_dict=fmt2)
        return


    def optimize_str_formatting(self, src):
        assert isinstance(src, list)
        # Concatenate styles
        s = []
        for x in src:
            if isinstance(x, dict) and len(s) > 0 and isinstance(s[-1], dict):
                s[-1].update(x)
            else:
                s.append(x)
        # Remove styles at the end of a sequence
        while len(s)>0 and isinstance(s[-1], dict):
            print "XLBUF Warning: sequence ends with style:", s
            del s[-1]
        i = 0
        a = []
        l = len(s)
        try:
            while(i < l):
                v = s[i]
                if isstr(v): # non-formatted string
                    while i+1<l and isstr(s[i+1]): # append following simple strings
                        v += s[i+1]
                        i += 1
                    a.append(v)
                    i += 1
                    continue
                # formatted string
                v = s[i+1]
                assert isstr(v)
                if v == "": # skip empty formatted strings
                    i += 2
                    continue
                # concatenate strings with same format
                while i+2<l and isdict(s[i+2]) and s[i] == s[i+2]:
                    v += s[i+3]
                    i += 2
                a += [s[i], v]
                i += 2
        except:
            print s
            raise
        return a

    def write_all(self, worksheet, out_cell_fmt):
        #ff = open('/tmp/test.txt', 'w')
        MAX_URLS = 65530
        for cell in self.buf.itervalues():
            fmt_dict = {}
            fmt_dict.update(cell.fmt_dict)
            if cell.y in self.row_format:
                fmt_dict.update(self.row_format[cell.y])
            fmt = self.get_xl_fmt(fmt_dict)
            out_cell_fmt[cell.ref] = fmt
            if cell.url is not None:
                self.num_urls += 1
                if self.num_urls == MAX_URLS:
                    print "Warning: exceeded %d URLs per sheet, ignoring the rest"%MAX_URLS
                if self.num_urls > MAX_URLS:
                    cell.url = None
            if cell.url is None:
                if isstr(cell.val) or cell.val is None:
                    worksheet.write(cell.y, cell.x, cell.val, fmt)
                else:
                    # if the value is an array, consider it a rich string with embedded formats
                    opt_vals = self.optimize_str_formatting(cell.val)
                    values = [x if isstr(x) else self.get_xl_fmt(x) for x in opt_vals]
                    values.append(fmt)
                    worksheet.write_rich_string(cell.y, cell.x, *values)
                    #ff.write("%s %s %s\n"%(cell.x, cell.y, values))
            else:
                worksheet.write_url(cell.y, cell.x, cell.url, fmt, cell.val)
            if cell.comment is not None:
                worksheet.write_comment(cell.y, cell.x, cell.comment, {'x_scale': 3.0, 'y_scale': 0.6})
