#!/usr/bin/python
# -*- coding UTF-8 -*-
# email hongling0@gmail.com

import traceback
import xlrd
import os
import math
import configparser
import sys
import time
import re
import tolua
import importlib
import traceback

importlib.reload(sys)

config = configparser.ConfigParser()
config['path'] = {
    'USEG': '1'
}
config.read_file(open("cfg.ini"))


def init_alise():
    ret = {}
    with open("alise.txt") as alise:
        for line in alise.readlines():
            if not line.startswith("#"):
                line = line.rstrip("\n")
                if line != "":
                    l = line.split("=")
                    ret[l[0]] = l[1].rstrip("\n")
    return ret


ALISE = init_alise()


def parser_integer(s, attr):
    if not isinstance(s, float):
        if s == "" and attr.find("e") != -1:
            return None
    return int(round(float(s)))


def parser_double(s, attr):
    if not isinstance(s, float):
        if s == "" and attr.find("e") != -1:
            return None
    if int(float(s)) == float(s):
        return int(float(s))
    return float(s)


def parser_string(s, attr):
    if s == "" and attr.find("e") != -1:
        return None
    return str(s)


def parser_luacode(s, attr):
    if s == "" and attr.find("e") != -1:
        return None
    return tolua.luacode(s)


PARSERLIST = {}
PARSERLIST["integer"] = parser_integer
PARSERLIST["double"] = parser_double
PARSERLIST["string"] = parser_string
PARSERLIST["luacode"] = parser_luacode


def getparser(ptype):
    parser = PARSERLIST.get(ptype, None)
    if parser:
        return parser
    alise = ALISE.get(ptype, None)
    if alise:
        parser = buildalise(alise)
        if parser:
            PARSERLIST[ptype] = parser
            return parser

    raise Exception("unknow type " + ptype.decode())


def buildalise(ptype):
    mobj = re.match(r"array<(.*?),(.*)>", ptype)
    if mobj:
        sp = mobj.group(1)
        subcall = []
        for v in str(mobj.group(2)).split(","):
            subcall.append(getparser(v))

        def array_func(s, attr):
            r = []
            if s == "":
                if attr.find("e") != -1:
                    return None
                else:
                    return r
            l = str(s).split(sp)
            for i in range(len(l)):
                if(i < len(subcall)):
                    r.append(subcall[i](l[i], attr))
                else:
                    r.append(subcall[-1](l[i], attr))
            return r
        return array_func


class rowctx:
    def __init__(self, owner):
        self.key_s = None
        self.key_c = None
        self.row_s = {}
        self.row_c = {}
        self.owner = owner
        self.flag = 3

    def readflag(self, coltype, attr, val):
        if attr == 'limit':
            self.flag = 0
            if val.find("c") != -1:
                self.flag |= (1 << 0)
            if val.find("s") != -1:
                self.flag |= (1 << 1)
            return True

    def read_ceil(self, coltype, colname, attr, val):
        if not self.readflag(coltype, attr, val):
            parser = getparser(coltype)
            val_s = parser(val, attr.replace("c", "") + "s")
            val_c = parser(val, attr.replace("s", "") + "c")
            if attr.find("k") != -1:
                if self.key_s != None:
                    raise Exception("mult key using")
                self.key_s = val_s

                if self.key_c != None:
                    raise Exception("mult key using")
                self.key_c = val_c

            if attr.find("s") != -1:
                self.row_s[colname] = val_s

            if attr.find("c") != -1:
                self.row_c[colname] = val_c

    def setvalue(self, k_coltype, k_attr, k_val, v_coltype, v_attr, v_val):
        k_parser = getparser(k_coltype)
        v_parser = getparser(v_coltype)

        if k_attr.find("s") != -1:
            self.key_s = k_parser(k_val, k_attr.replace("c", "") + "s")
            self.row_s = v_parser(v_val, v_attr.replace("c", "") + "s")

        if k_attr.find("c") != -1:
            self.key_c = k_parser(k_val, k_attr.replace("s", "") + "c")
            self.row_c = v_parser(v_val, v_attr.replace("s", "") + "c")

    def finish(self):
        if self.key_c and (self.flag & (1 << 0)):
            self.owner.change_c(self.key_c, self.row_c)
        if self.key_s and (self.flag & (1 << 1)):
            self.owner.change_s(self.key_s, self.row_s)


class tablectx:
    def __init__(self, owner, name):
        self.owner = owner
        self.name = name
        self.table_s = None
        self.table_c = None

    def change_s(self, key, row):
        if self.table_s == None:
            self.table_s = {}
        self.table_s[key] = row

    def change_c(self, key, row):
        if self.table_c == None:
            self.table_c = {}
        self.table_c[key] = row

    def finish(self):
        if self.table_s:
            self.owner.change_s(self.name, self.table_s)
        if self.table_c:
            self.owner.change_c(self.name, self.table_c)


class sheetctx:
    def __init__(self, name):
        self.name = {}
        self.table_s = None
        self.table_c = None

    def change_s(self, key, row):
        if self.table_s == None:
            self.table_s = {}
        if self.table_s.get(key, None):
            raise Exception("dumplicate global index" + key)
        self.table_s[key] = row

    def change_c(self, key, row):
        if self.table_c == None:
            self.table_c = {}
        if self.table_c.get(key, None):
            raise Exception("dumplicate global index" + key)
        self.table_c[key] = row


def transfer_z(sctx, bootsheet):
    name = bootsheet.name
    if bootsheet.nrows < 4:
        raise Exception("Error format " + name)

    tctx = tablectx(sctx, name[2:])
    for row in range(bootsheet.nrows):
        if(row < 4):
            continue
        else:
            rctx = rowctx(tctx)
            skip = False
            for col in range(bootsheet.ncols):
                coltype = bootsheet.cell(1, col).value
                colname = bootsheet.cell(2, col).value
                colattr = bootsheet.cell(3, col).value
                cellval = bootsheet.cell(row, col).value

                if str(cellval).startswith("//"):
                    skip = True
                    break

                if len(colattr) == 0:
                    continue
                # rctx.read_ceil(coltype, colname, colattr, cellval)
                try:
                    rctx.read_ceil(coltype, colname, colattr, cellval)
                except Exception as e:
                    raise Exception("Exception @" + name + "." + bootsheet.name
                                    + ("(") + str(row+1) + ", " + str(col) + ")\n"
                                    + repr(e) + "\n"
                                    + traceback.format_exc())
            if not skip:
                rctx.finish()

    tctx.finish()


def transfer_y(sctx, bootsheet):
    name = bootsheet.name
    if bootsheet.nrows < 3:
        raise Exception("Error format " + name)

    k_coltype = bootsheet.cell(1, 0).value
    k_colattr = bootsheet.cell(2, 0).value
    v_coltype = bootsheet.cell(1, 1).value
    v_colattr = bootsheet.cell(2, 1).value
    l_coltype = bootsheet.cell(1, 2).value
    l_colattr = bootsheet.cell(2, 2).value

    tctx = tablectx(sctx, name[2:])
    for row in range(bootsheet.nrows):
        if(row < 3):
            continue
        else:
            rctx = rowctx(tctx)

            k_cellval = bootsheet.cell(row, 0).value
            if str(k_cellval).startswith("//"):
                continue
            v_cellval = bootsheet.cell(row, 1).value
            if str(v_cellval).startswith("//"):
                continue

            try:
                rctx.readflag(l_coltype, l_colattr,
                              bootsheet.cell(row, 2).value)
                rctx.setvalue(k_coltype, k_colattr, k_cellval,
                              v_coltype, v_colattr, v_cellval)
            except Exception as e:
                raise Exception("Exception @"+name
                                + "." + bootsheet.name
                                + ("(") + str(row + 1) + ")\n"
                                + repr(e) + "\n"
                                + traceback.format_exc())

            rctx.finish()

    tctx.finish()


def transfer_g(sctx, bootsheet):
    name = bootsheet.name
    if bootsheet.nrows < 3:
        raise Exception("Error format " + name)

    tctx = tablectx(sctx, "_G")
    for row in range(bootsheet.nrows):
        if(row < 3):
            continue
        else:
            rctx = rowctx(tctx)

            k_coltype = bootsheet.cell(1, 0).value
            k_colattr = bootsheet.cell(2, 0).value
            k_cellval = bootsheet.cell(row, 0).value
            if str(k_cellval).startswith("//"):
                continue

            v_coltype = bootsheet.cell(1, 1).value
            v_colattr = bootsheet.cell(2, 1).value
            v_cellval = bootsheet.cell(row, 1).value
            if str(v_cellval).startswith("//"):
                continue

            try:
                rctx.setvalue(k_coltype, k_colattr, k_cellval,
                              v_coltype, v_colattr, v_cellval)
            except Exception as e:
                import traceback
                raise Exception("Exception @" + name + "." + bootsheet.name
                                + ("(") + str(row + 1) + ")\n"
                                + repr(e) + "\n"
                                + traceback.format_exc())

            rctx.finish()

    tctx.finish()


def transferfile(name, path):
    sctx = sheetctx(name)
    workbook = xlrd.open_workbook(path)
    for booksheet in workbook.sheets():
        booksheetname = booksheet.name
        if booksheetname.startswith('z_'):
            transfer_z(sctx, booksheet)
            continue
        if booksheetname.startswith("y_"):
            transfer_y(sctx, booksheet)
            continue

        if booksheetname == "_G":
            transfer_g(sctx, booksheet)
            continue

    return sctx


def readxlsx(indir):
    ret = {}
    fs = os.listdir(indir)
    for f in fs:
        fname = indir+"/"+f
        (name, ext) = os.path.splitext(f)
        if os.path.isfile(fname) and (fname.endswith(".xlsx") or fname.endswith('.xls')):
            ret[name] = fname
    return ret


def makexlsxlist():
    xls_list = {}
    for f in sys.argv[1:]:
        (fpath, tmp) = os.path.split(f)
        (fname, ext) = os.path.splitext(tmp)
        if f.endswith('.xlsx') or f.endswith('.xls'):
            xls_list[fname] = f
    return xls_list


def eacho_tables(tables, cb):
    def eacho_tables_inner(datas, show):
        keys = sorted(datas.keys())
        for f in keys:
            if show:
                print("\tadd "+f)
            cb(f, datas[f])

    data = tables.get("_G", None)
    if data:
        print("\tadd _G")
        tables.pop("_G")
        eacho_tables_inner(data, False)

    eacho_tables_inner(tables, True)


def abspath(path):
    if os.path.isabs(path):
        return path
    return os.path.abspath(path)


def trans2lua(sctx, name, path_s, path_c):
    deep = config.getint("path", "DEEP")
    if sctx.table_s:
        table = sctx.table_s
        fname = os.path.join(path_s, name + ".lua")
        print(fname)
        with open(fname, 'w', encoding='UTF-8') as out:
            if 1 == config.getint("path", "USEG"):
                def writer_s(f, data):
                    out.write(f)
                    out.write(" = ")
                    out.write(tolua.trans_obj(data, 0, deep).encode('utf8'))
                    out.write("\n")
                eacho_tables(table, writer_s)
            else:
                out.write("return ")
                out.write(tolua.trans_obj(table, 0, deep + 1))
                for f in sorted(table.keys()):
                    print("\tadd "+f)
                out.write("\n")
            out.close()

    if sctx.table_c:
        table = sctx.table_c
        fname = os.path.join(path_c, "prop_" + name + ".lua")
        print(fname)
        with open(fname, 'w', encoding='UTF-8') as out:
            out.write("module(\"resmng\")\n\n")

            def writer_c(f, data):
                f_name = 'prop%s' % (f.capitalize())
                f_name_data = f_name + 'Data'
                out.write(f_name_data)
                out.write(" = ")
                out.write(tolua.trans_obj(data, 0, deep))
                out.write("\n\n")

            eacho_tables(table, writer_c)


def main(xls_list):
    fs = xls_list
    if len(fs) == 0:
        fs = readxlsx(config.get("path", "IN"))
    path_s = config.get("path", "SERVER_OUT")
    path_c = config.get("path", "CLINET_OUT")
    if not os.path.exists(path_s):
        os.makedirs(path_s)
    if not os.path.exists(path_c):
        os.makedirs(path_c)

    for f in fs.keys():
        path = fs[f]
        print("transferfile " + path)
        sctx = transferfile(f, path)

        trans2lua(sctx, f, abspath(path_s), abspath(path_c))


if __name__ == "__main__":
    try:
        main(makexlsxlist())
    except Exception as e:
        import traceback
        msg = traceback.format_exc()
        print(e)
        print(msg)
        print(e.args)
    finally:
        input("Press enter to quit")
