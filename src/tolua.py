#!/usr/bin/python
# -*- coding UTF-8 -*-
# email hongling0@gmail.com
import os


class luacode:
    def __init__(self, s):
        if isinstance(s, float):
            if int(s) == s:
                s = int(s)
        elif s == "":
            s = "\"\""
        self._str = str(s)

    def __str__(self):
        return self._str


def septer(deep, ending_deep):
    if deep > ending_deep:
        return "", "", " "
    else:
        return "    ", "\n", "\n"


def multsp(n, sp):
    return sp * n


def trans_dict(obj, deep, ending_deep):
    numchar = ("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    space, ending, sep = septer(deep, ending_deep)
    keys = sorted(obj.keys())
    vals = []
    for k in keys:
        val = obj[k]
        if val is not None:
            v = multsp(deep, space)
            if isinstance(k, int):
                v = v + "[" + str(k) + "] = "
            elif isinstance(k, str):
                try:
                    float(k)
                except ValueError as e:
                    if k.startswith(numchar) or k.find("%") != -1:
                        v = v + "['" + str(k) + "'] = "
                    else:
                        v = v + str(k) + " = "
                else:
                    v = v + "['" + str(k) + "'] = "
            v = v + trans_obj(val, deep, ending_deep)
            vals.append(v)
            
    if len(vals)==0:
        ending=""

    return "{" + ending + ("," + sep).join(vals) + \
        ending + multsp(deep - 1, space) + "}"


def trans_list(obj, deep, ending_deep):
    space, ending, sep = septer(deep, ending_deep)
    vals = []
    for i in range(len(obj)):
        v = multsp(deep, space) + trans_obj(obj[i], deep, ending_deep)
        vals.append(v)
    if len(vals)==0:
        ending=""
    return "{" + ending + ("," + sep).join(vals) + \
        ending + multsp(deep - 1, space) + "}"


def trans_obj(obj, deep, ending_deep):
    if isinstance(obj, luacode):
        return str(obj)
    elif isinstance(obj, int):
        return str(obj)
    elif isinstance(obj, float):
        return str(obj)
    elif isinstance(obj, str):
        return "\"" + obj.replace("\"", "\\\"") + "\""
    elif isinstance(obj, list):
        return trans_list(obj, deep + 1, ending_deep)
    elif isinstance(obj, dict):
        return trans_dict(obj, deep + 1, ending_deep)
    elif obj is None:
        return "nil"
    else:
        raise Exception("Invalid obj " + str(type(obj)))


def lua_test(files):
    for f in files:
        r = os.popen("lua " + f)
