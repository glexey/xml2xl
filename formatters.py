import re

def xml_strip_formatter(s):

    s = re.sub(ur'^\s+|^\s*$|\s*$', u'', s, flags=re.M)
    s = re.sub(ur'\s+|\n', u' ', s, flags=re.M)

    return s

def fuse_formatter(s):

    s = xml_strip_formatter(s)
    s = re.sub(ur';\s*', ur'\n', s, flags=re.M)
    s = s.strip()
    return s

def make_hex_formatter(digits):
    def hex_formatter(s):
        try:
            n = to_int(s)
        except:
            return s
        fmt = "0x%%0%sx"%digits
        return fmt%n
    return hex_formatter

by_name = {
        'fuse_formatter': fuse_formatter,
        'hex': make_hex_formatter(''),
        'hex1': make_hex_formatter(1),
        'hex2': make_hex_formatter(2),
        'hex3': make_hex_formatter(3),
        'hex4': make_hex_formatter(4),
        'hex5': make_hex_formatter(5),
        'hex6': make_hex_formatter(6),
        }
