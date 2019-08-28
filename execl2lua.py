#!/usr/bin/python
# -*- coding UTF-8 -*-
# email hongling0@gmail.com

import xlrd
import os
import math
import ConfigParser
import sys
import time
import re
import tolua

reload(sys)
encoding = "utf-8"
sys.setdefaultencoding(encoding)

config = ConfigParser.ConfigParser()
config.readfp(open("cfg.ini"))


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
    return float(s)


def parser_string(s, attr):
    if not isinstance(s, float):
        if s == "" and attr.find("e") != -1:
            return None
    return str(s).encode(encoding)


def parser_any(s, attr):
    mobj = re.math(r"(.*?)\|(.*)", s)
    if mobj:
        parser = mobj.group(1)
        sn = mobj.group(2)
        return getparser(parser)(sn, attr)


PARSERLIST = {}
PARSERLIST["integer"] = parser_integer
PARSERLIST["double"] = parser_double
PARSERLIST["string"] = parser_string
PARSERLIST["any"] = parser_any


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

    raise Exception("unknow type " + ptype)


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
            for i in xrange(len(l)):
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

    def read_ceil(self, coltype, colname, attr, val):
        parser = getparser(coltype)
        val_s = parser(val, attr.replace("c", "")+"s")
        val_c = parser(val, attr.replace("s", "")+"c")
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

    def finish(self):
        if self.key_s:
            self.owner.change_s(self.key_s, self.row_s)
        if self.key_c:
            self.owner.change_c(self.key_c, self.row_c)


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
        self.table_s[key] = row

    def change_c(self, key, row):
        if self.table_c == None:
            self.table_c = {}
        self.table_c[key] = row


def transfer_z(sctx, bootsheet):
    name = bootsheet.name.encode(encoding)
    if bootsheet.nrows < 4:
        raise Exception("Error format " + name)

    tctx = tablectx(sctx, name[2:])
    for row in xrange(bootsheet.nrows):
        if(row < 4):
            continue
        else:
            rctx = rowctx(tctx)
            for col in xrange(bootsheet.ncols):
                coltype = bootsheet.cell(1, col).value.encode(encoding)
                colname = bootsheet.cell(2, col).value.encode(encoding)
                colattr = bootsheet.cell(3, col).value.encode(encoding)
                cellval = bootsheet.cell(row, col).value

                if len(colattr) == 0:
                    continue
                rctx.read_ceil(coltype, colname, colattr, cellval)
                try:
                    pass  # rctx.read_ceil(coltype, colname, colattr, cellval)
                except Exception as e:
                    import traceback
                    raise Exception("Exception @"+name+"."+bootsheet.name.encode(encoding)
                                    + ("(")+str(row)+", "+str(col)+")\n"
                                    + repr(e)+"\n"
                                    + traceback.format_exc())

            rctx.finish()

    tctx.finish()


def transferfile(name, path):
    sctx = sheetctx(name)
    workbook = xlrd.open_workbook(path)
    for booksheet in workbook.sheets():
        booksheetname = booksheet.name.encode(encoding)
        if booksheetname.startswith("z_"):
            transfer_z(sctx, booksheet)
            continue

    return sctx


def readxlsx(indir):
    ret = {}
    fs = os.listdir(indir)
    for f in fs:
        fname = indir+"/"+f
        (name, ext) = os.path.splitext(f)
        if os.path.isfile(fname):
            ret[name] = fname
    return ret


def trans2lua(sctx, name):
    deep = int(config.get("path", "DEEP"))

    if sctx.table_s:
        table = sctx.table_s
        fname = config.get("path", "SERVER_OUT") + "/" + name + ".lua"
        print("\t"+fname)
        out = open(fname, "w")
        keys = table.keys()
        keys.sort()
        for f in keys:
            print("\t\tadd " + f)
            out.write(f)
            out.write("=")
            out.write(tolua.trans_obj(table[f], 0, deep))
            out.write("\n")

        out.close()

    if sctx.table_c:
        table = sctx.table_c
        fname = config.get("path", "CLINET_OUT") + "/" + name + ".lua"
        print("\t"+fname)
        out = open(fname, "w")
        out.write("module(\"" + name + "\")\n")
        keys = table.keys()
        keys.sort()
        for f in keys:
            print("\t\tadd " + f)
            out.write(f)
            out.write("=")
            out.write(tolua.trans_obj(table[f], 0, deep))
            out.write("\n")

        out.close()


def main():
    fs = readxlsx(config.get("path", "IN"))

    for f in fs.keys():
        path = fs[f]
        print("transferfile " + path)
        sctx = transferfile(f, path)

        trans2lua(sctx, f)

    print("Press enter to quit")
    input()


if __name__ == "__main__":
    main()