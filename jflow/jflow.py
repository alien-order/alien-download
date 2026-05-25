#!/usr/bin/env python3
"""
jflow - Spring MVC + MyBatis + JSP process-flow extractor (no AI, pure static analysis).

Point it at a Java web-app source tree. For a chosen entry point (a URL, a JSP, or a
controller method) it traces the call flow:

    JSP url  ->  @Controller method  ->  @Service impl  ->  @Mapper  ->  SQL -> tables

How it resolves the chain (all deterministic, no AI):
  * URL          : class-level + method-level @RequestMapping/@GetMapping/... values.
  * field calls  : `someService.foo()` -> declared field type -> bean type.
  * interface    : interface type -> implementing @Service/@Repository/@Component class.
  * mapper       : @Mapper interface method == MyBatis XML <select|insert|...> id
                   (XML namespace == interface FQN). Also @Select/@Insert annotations.
  * sql -> table : FROM / JOIN / INTO / UPDATE parsing (with <include refid> expansion).

Usage:
    python jflow.py stats --src PATH [--src PATH2 ...]
    python jflow.py urls  --src PATH [...] [--context-path /app] [--grep order]
    python jflow.py flow  --src PATH [...] --url    /order/save     [--mermaid]
    python jflow.py flow  --src PATH [...] --method com.x.OrderController.save
    python jflow.py flow  --src PATH [...] --jsp    webapp/order/list.jsp

Dependency:  pip install javalang
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Optional
import xml.etree.ElementTree as ET

try:
    import javalang
    import javalang.tree as JT
except ImportError:
    sys.stderr.write(
        "ERROR: javalang is not installed.\n"
        "       Run:  pip install javalang\n"
    )
    sys.exit(2)


# --------------------------------------------------------------------------- #
# Model
# --------------------------------------------------------------------------- #
@dataclass
class MethodCall:
    qualifier: Optional[str]   # textual receiver: "orderService", "this.x", or None (same class)
    member: str                # called method name


@dataclass
class JavaMethod:
    name: str
    annotations: dict = field(default_factory=dict)   # simple ann name -> [str values]
    http_methods: list = field(default_factory=list)  # GET/POST/...
    url_paths: list = field(default_factory=list)      # method-level mapping paths
    calls: list = field(default_factory=list)          # [MethodCall]
    return_views: list = field(default_factory=list)   # returned view-name string literals


@dataclass
class JavaType:
    fqn: str
    name: str
    package: str
    kind: str                                  # 'class' | 'interface' | 'enum'
    annotations: dict = field(default_factory=dict)
    imports: dict = field(default_factory=dict)        # simple name -> fqn
    extends: Optional[str] = None
    implements: list = field(default_factory=list)     # [simple names]
    fields: dict = field(default_factory=dict)         # field name -> type simple name
    methods: dict = field(default_factory=dict)        # method name -> JavaMethod
    class_url_paths: list = field(default_factory=list)
    source_file: str = ""

    def has(self, *names):
        return any(n in self.annotations for n in names)

    def is_controller(self):
        return self.has("Controller", "RestController")

    def is_service(self):
        return self.has("Service")

    def is_repository(self):
        return self.has("Repository")

    def is_component(self):
        return self.has("Service", "Repository", "Component")


@dataclass
class MapperStatement:
    namespace: str
    id: str
    stmt_type: str            # select/insert/update/delete
    tables: list
    sql_preview: str
    source_file: str


@dataclass
class JspPage:
    path: str
    urls: list = field(default_factory=list)
    includes: list = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Java parsing
# --------------------------------------------------------------------------- #
MAPPING_ANN = {
    "RequestMapping": None,
    "GetMapping": "GET", "PostMapping": "POST", "PutMapping": "PUT",
    "DeleteMapping": "DELETE", "PatchMapping": "PATCH",
}
SQL_ANN = {"Select": "select", "Insert": "insert", "Update": "update", "Delete": "delete"}


def _ann_simple(ann):
    return ann.name.split(".")[-1]


def _unquote(v):
    if v and len(v) >= 2 and v[0] in "\"'" and v[-1] in "\"'":
        return v[1:-1]
    return v


def _ann_string_values(ann):
    """String values of value=/path= (or the single unnamed value) of an annotation."""
    out = []

    def collect(node):
        if node is None:
            return
        if isinstance(node, list):
            for x in node:
                collect(x)
        elif isinstance(node, JT.ElementArrayValue):
            for x in node.values:
                collect(x)
        elif isinstance(node, JT.Literal):
            out.append(_unquote(node.value))

    el = ann.element
    if isinstance(el, list):
        for item in el:
            if isinstance(item, JT.ElementValuePair):
                if item.name in ("value", "path"):
                    collect(item.value)
            else:
                collect(item)
    else:
        collect(el)
    return out


def _type_name(t):
    if t is None:
        return None
    return getattr(t, "name", None)


def _parse_java_file(path):
    types = []
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            src = f.read()
        tree = javalang.parse.parse(src)
    except Exception as e:                       # JavaSyntaxError, LexerError, ...
        return types, "%s: %s" % (type(e).__name__, e)

    package = tree.package.name if tree.package else ""
    imports = {}
    for imp in (tree.imports or []):
        if imp.wildcard or imp.static:
            continue
        imports[imp.path.split(".")[-1]] = imp.path

    for t in tree.types:
        if not isinstance(t, (JT.ClassDeclaration, JT.InterfaceDeclaration)):
            continue
        kind = "interface" if isinstance(t, JT.InterfaceDeclaration) else "class"
        fqn = (package + "." + t.name) if package else t.name

        anns = {}
        class_urls = []
        for a in (t.annotations or []):
            name = _ann_simple(a)
            anns[name] = _ann_string_values(a)
            if name == "RequestMapping":
                class_urls = _ann_string_values(a) or [""]

        implements = [_type_name(it) for it in (getattr(t, "implements", None) or []) if _type_name(it)]
        extends = None
        ext = getattr(t, "extends", None)
        if isinstance(ext, list):
            extends = _type_name(ext[0]) if ext else None
        elif ext is not None:
            extends = _type_name(ext)

        fields = {}
        for fld in t.fields:
            tn = _type_name(fld.type)
            if not tn:
                continue
            for d in fld.declarators:
                fields[d.name] = tn

        methods = {}
        for m in t.methods:
            jm = methods.get(m.name)
            if jm is None:
                jm = JavaMethod(name=m.name)
                methods[m.name] = jm
            for a in (m.annotations or []):
                an = _ann_simple(a)
                jm.annotations[an] = _ann_string_values(a)
                if an in MAPPING_ANN:
                    jm.url_paths.extend(_ann_string_values(a) or [""])
                    if MAPPING_ANN[an]:
                        jm.http_methods.append(MAPPING_ANN[an])
            if m.body:
                try:
                    for _p, inv in m.filter(JT.MethodInvocation):
                        jm.calls.append(MethodCall(qualifier=inv.qualifier, member=inv.member))
                except Exception:
                    pass
                try:
                    for _p, ret in m.filter(JT.ReturnStatement):
                        ex = ret.expression
                        if isinstance(ex, JT.Literal):
                            s = _unquote(ex.value)
                            if s and "/" not in s[:1] and "<" not in s and len(s) < 80:
                                jm.return_views.append(s)
                        elif isinstance(ex, JT.ClassCreator) and _type_name(ex.type) == "ModelAndView":
                            if ex.arguments and isinstance(ex.arguments[0], JT.Literal):
                                jm.return_views.append(_unquote(ex.arguments[0].value))
                except Exception:
                    pass

        types.append(JavaType(
            fqn=fqn, name=t.name, package=package, kind=kind, annotations=anns,
            imports=imports, extends=extends, implements=implements, fields=fields,
            methods=methods, class_url_paths=class_urls, source_file=path))
    return types, None


# --------------------------------------------------------------------------- #
# MyBatis mapper XML parsing
# --------------------------------------------------------------------------- #
_TABLE_RE = re.compile(
    r"\b(?:from|join|into|update)\s+([A-Za-z_][\w$]*(?:\.[A-Za-z_][\w$]*)?)", re.I)
_NOISE_TABLE = {"select", "where", "set", "values", "dual", "on", "and", "or"}


def _clean_sql(sql):
    sql = re.sub(r"#\{[^}]*\}", "?", sql)
    sql = re.sub(r"\$\{[^}]*\}", "?", sql)
    return sql


def _extract_tables(sql):
    tabs = []
    for m in _TABLE_RE.finditer(_clean_sql(sql)):
        t = m.group(1)
        if t.lower() in _NOISE_TABLE:
            continue
        if t not in tabs:
            tabs.append(t)
    return tabs


def _regex_mapper(path, text):
    ns = re.search(r"<mapper\s+namespace\s*=\s*[\"']([^\"']+)[\"']", text, re.I)
    if not ns:
        return [], None
    namespace = ns.group(1)
    stmts = []
    for m in re.finditer(
            r"<(select|insert|update|delete)\b[^>]*\bid\s*=\s*[\"']([^\"']+)[\"'][^>]*>(.*?)</\1>",
            text, re.I | re.S):
        stype, sid, body = m.group(1).lower(), m.group(2), m.group(3)
        body = re.sub(r"<[^>]+>", " ", body)
        stmts.append(MapperStatement(namespace, sid, stype, _extract_tables(body),
                                     re.sub(r"\s+", " ", body).strip()[:300], path))
    return stmts, namespace


def _parse_mapper_xml(path):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
    except Exception:
        return [], None
    if "<mapper" not in text or "namespace" not in text:
        return [], None
    try:
        root = ET.fromstring(text)
    except Exception:
        return _regex_mapper(path, text)
    if root.tag != "mapper":
        return [], None
    namespace = root.attrib.get("namespace")
    if not namespace:
        return [], None

    # <sql id=...> fragments for <include refid=...> expansion
    fragments = {}
    for el in root.iter("sql"):
        fid = el.attrib.get("id")
        if fid:
            fragments[fid] = "".join(el.itertext())

    def expand(el):
        parts = [el.text or ""]
        for child in el:
            if child.tag == "include":
                ref = child.attrib.get("refid", "")
                ref = ref.split(".")[-1] if "." in ref else ref
                parts.append(fragments.get(ref, ""))
            else:
                parts.append(expand(child))
            parts.append(child.tail or "")
        return " ".join(parts)

    stmts = []
    for el in root:
        if el.tag in ("select", "insert", "update", "delete"):
            sid = el.attrib.get("id")
            if not sid:
                continue
            sql = expand(el)
            stmts.append(MapperStatement(
                namespace, sid, el.tag, _extract_tables(sql),
                re.sub(r"\s+", " ", sql).strip()[:300], path))
    return stmts, namespace


# --------------------------------------------------------------------------- #
# JSP parsing
# --------------------------------------------------------------------------- #
_JSP_URL_PATS = [
    re.compile(r"action\s*=\s*[\"']([^\"']+)[\"']", re.I),
    re.compile(r"href\s*=\s*[\"']([^\"']+)[\"']", re.I),
    re.compile(r"\burl\s*:\s*[\"']([^\"']+)[\"']", re.I),
    re.compile(r"location\.href\s*=\s*[\"']([^\"']+)[\"']", re.I),
    re.compile(r"<c:url\s+value\s*=\s*[\"']([^\"']+)[\"']", re.I),
]
_JSP_INCLUDE_PATS = [
    re.compile(r"<%@\s*include\s+file\s*=\s*[\"']([^\"']+)[\"']", re.I),
    re.compile(r"<jsp:include\s+page\s*=\s*[\"']([^\"']+)[\"']", re.I),
]


def _parse_jsp(path):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            txt = f.read()
    except Exception:
        return JspPage(path)
    urls, incs = [], []
    for pat in _JSP_URL_PATS:
        urls += [m.group(1).strip() for m in pat.finditer(txt)]
    for pat in _JSP_INCLUDE_PATS:
        incs += [m.group(1).strip() for m in pat.finditer(txt)]
    return JspPage(path, list(dict.fromkeys(urls)), list(dict.fromkeys(incs)))


# --------------------------------------------------------------------------- #
# Index + resolver
# --------------------------------------------------------------------------- #
SKIP_DIRS = {".git", ".svn", ".hg", "target", "build", "out", "bin",
             "node_modules", ".idea", ".gradle", ".settings"}


def _join_url(*parts):
    segs = [p.strip("/") for p in parts if p]
    return "/" + "/".join(s for s in segs if s)


def _norm(u):
    return "/" + u.strip("/") if u else "/"


def _pattern_match(pattern, url):
    ps, us = _norm(pattern).strip("/").split("/"), _norm(url).strip("/").split("/")
    i = 0
    for i, seg in enumerate(ps):
        if seg == "**":
            return True
        if i >= len(us):
            return False
        if seg.startswith("{") and seg.endswith("}"):
            continue
        if seg in ("*",):
            continue
        if seg != us[i]:
            # allow Spring path-extension/regex sloppiness
            if "{" in seg and "}" in seg:
                continue
            return False
    return len(ps) == len(us)


class Index:
    def __init__(self):
        self.types: dict[str, JavaType] = {}
        self.simple: dict[str, set] = {}
        self.mapper_stmts: dict[tuple, MapperStatement] = {}
        self.mapper_ns: set = set()
        self.jsp: dict[str, JspPage] = {}
        self.url_map: list = []          # (url, [http], controller_fqn, method)
        self.errors: list = []
        self.n_java = self.n_xml = self.n_jsp = 0

    # ----- build --------------------------------------------------------- #
    def add_type(self, t: JavaType):
        self.types[t.fqn] = t
        self.simple.setdefault(t.name, set()).add(t.fqn)

    def build_url_map(self):
        self.url_map = []
        for t in self.types.values():
            if not t.is_controller():
                continue
            bases = t.class_url_paths or [""]
            for m in t.methods.values():
                if not m.url_paths:
                    continue
                for b in bases:
                    for mp in m.url_paths:
                        self.url_map.append((_join_url(b, mp), m.http_methods, t.fqn, m.name))

    # ----- type resolution ---------------------------------------------- #
    def resolve_type_name(self, owner: JavaType, simple):
        if not simple:
            return None
        fqn = owner.imports.get(simple)
        if fqn and fqn in self.types:
            return fqn
        cand = (owner.package + "." + simple) if owner.package else simple
        if cand in self.types:
            return cand
        s = self.simple.get(simple)
        if s and len(s) == 1:
            return next(iter(s))
        return fqn  # external (FQN known but not scanned) or None

    def resolve_field(self, owner: JavaType, field_name):
        return self.resolve_type_name(owner, owner.fields.get(field_name))

    def resolve_impl(self, iface_fqn):
        t = self.types.get(iface_fqn)
        if t is None or t.kind == "class":
            return iface_fqn
        impls = []
        for ot in self.types.values():
            if ot.kind != "class":
                continue
            for inm in ot.implements:
                if self.resolve_type_name(ot, inm) == iface_fqn:
                    impls.append(ot.fqn)
                    break
        comp = [f for f in impls if self.types[f].is_component()]
        if comp:
            return comp[0]
        if impls:
            return impls[0]
        return iface_fqn   # likely a MyBatis mapper interface (proxy, no source impl)

    def kind_of(self, t: Optional[JavaType]):
        if t is None:
            return "external"
        if t.fqn in self.mapper_ns or t.has("Mapper"):
            return "mapper"
        if t.is_controller():
            return "controller"
        if t.is_service():
            return "service"
        if t.is_repository():
            return "dao"
        if t.has("Component"):
            return "component"
        return "class"

    # ----- url matching -------------------------------------------------- #
    def normalize_jsp_url(self, u, context_path=""):
        u = u.strip().split("?", 1)[0].split("#", 1)[0]
        u = re.sub(r"^\$\{[^}]*\}", "", u)        # ${ctx}
        u = re.sub(r"^<%=.*?%>", "", u)
        u = re.sub(r"<c:out[^>]*>.*?</c:out>", "", u)
        if context_path and u.startswith(context_path):
            u = u[len(context_path):]
        if not u or u.startswith(("javascript:", "mailto:", "#", "http://", "https://")):
            return None
        return _norm(u)

    def match_url(self, url):
        url = _norm(url)
        exact = [e for e in self.url_map if _norm(e[0]) == url]
        if exact:
            return exact
        return [e for e in self.url_map if _pattern_match(e[0], url)]


def build_index(roots, verbose=False):
    idx = Index()
    java, xml, jsp = [], [], []
    for root in roots:
        for dp, dn, fn in os.walk(root):
            dn[:] = [d for d in dn if d not in SKIP_DIRS]
            for f in fn:
                p = os.path.join(dp, f)
                low = f.lower()
                if low.endswith(".java"):
                    java.append(p)
                elif low.endswith(".xml"):
                    xml.append(p)
                elif low.endswith((".jsp", ".jspf")):
                    jsp.append(p)
    idx.n_java, idx.n_xml, idx.n_jsp = len(java), len(xml), len(jsp)

    for p in java:
        types, err = _parse_java_file(p)
        if err:
            idx.errors.append((p, err))
        for t in types:
            idx.add_type(t)
    for p in xml:
        stmts, ns = _parse_mapper_xml(p)
        if ns:
            idx.mapper_ns.add(ns)
            for s in stmts:
                idx.mapper_stmts[(s.namespace, s.id)] = s
    for p in jsp:
        idx.jsp[p] = _parse_jsp(p)
    idx.build_url_map()
    if verbose:
        sys.stderr.write(
            "[jflow] %d java types, %d mapper namespaces, %d url mappings, "
            "%d jsp, %d parse errors\n" % (len(idx.types), len(idx.mapper_ns),
                                           len(idx.url_map), len(idx.jsp), len(idx.errors)))
    return idx


# --------------------------------------------------------------------------- #
# Flow tracing
# --------------------------------------------------------------------------- #
@dataclass
class FlowNode:
    kind: str
    label: str
    note: str = ""
    children: list = field(default_factory=list)


def _field_name_of(qualifier):
    """Map a call qualifier to a single field name, or None if unresolvable."""
    if qualifier is None:
        return None
    q = qualifier
    if q == "this":
        return None
    if q.startswith("this.") and q.count(".") == 1:
        return q[5:]
    if "." in q:
        return None      # chained / package-qualified: skip
    return q


def trace_method(idx: Index, owner_fqn, method_name, visited=None, depth=0, max_depth=30):
    visited = visited if visited is not None else set()
    owner = idx.types.get(owner_fqn)
    kind = idx.kind_of(owner)
    short = owner.name if owner else owner_fqn.split(".")[-1]
    node = FlowNode(kind, "%s.%s" % (short, method_name))

    key = owner_fqn + "#" + method_name
    if key in visited:
        node.note = "(already traced above)"
        return node
    visited.add(key)
    if depth > max_depth:
        node.note = "(max depth)"
        return node

    # ---- mapper terminal: method -> SQL ---- #
    if kind == "mapper":
        st = idx.mapper_stmts.get((owner_fqn, method_name))
        if st:
            tables = ", ".join(st.tables) if st.tables else "?"
            node.children.append(FlowNode(
                "sql", "%s -> %s" % (st.stmt_type.upper(), tables), st.sql_preview))
            return node
        jm = owner.methods.get(method_name) if owner else None
        if jm:
            for an, stype in SQL_ANN.items():
                if an in jm.annotations and jm.annotations[an]:
                    sql = " ".join(jm.annotations[an])
                    tables = ", ".join(_extract_tables(sql)) or "?"
                    node.children.append(FlowNode(
                        "sql", "%s -> %s" % (stype.upper(), tables),
                        re.sub(r"\s+", " ", sql)[:300]))
                    return node
        node.note = "(no mapped statement found)"
        return node

    if owner is None:
        node.kind = "external"
        node.note = "(source not found)"
        return node

    jm = owner.methods.get(method_name)
    if jm is None:
        node.note = "(method not found in source)"
        return node

    # views rendered by a controller method
    for v in dict.fromkeys(jm.return_views):
        node.children.append(FlowNode("view", "view: %s" % v))

    seen_edges = set()
    for call in jm.calls:
        fname = _field_name_of(call.qualifier)
        if fname is None:
            # same-class unqualified helper call -> follow only if it exists here
            if call.qualifier is None and call.member in owner.methods and call.member != method_name:
                edge = (owner_fqn, call.member)
                if edge in seen_edges:
                    continue
                seen_edges.add(edge)
                node.children.append(trace_method(idx, owner_fqn, call.member, visited, depth + 1, max_depth))
            continue
        target_fqn = idx.resolve_field(owner, fname)
        if not target_fqn or target_fqn not in idx.types:
            continue                       # local var, param, or external library: skip
        target_fqn = idx.resolve_impl(target_fqn)
        edge = (target_fqn, call.member)
        if edge in seen_edges:
            continue
        seen_edges.add(edge)
        node.children.append(trace_method(idx, target_fqn, call.member, visited, depth + 1, max_depth))
    return node


def trace_url(idx: Index, url):
    matches = idx.match_url(url)
    roots = []
    for (mapped, https, cfqn, mname) in matches:
        head = FlowNode("url", "%s %s" % ("/".join(https) if https else "ANY", mapped))
        head.children.append(trace_method(idx, cfqn, mname, set()))
        roots.append(head)
    return roots, matches


# --------------------------------------------------------------------------- #
# Rendering
# --------------------------------------------------------------------------- #
TAG = {
    "url": "[URL]", "controller": "[CTRL]", "service": "[SVC]", "dao": "[DAO]",
    "mapper": "[MAP]", "sql": "[SQL]", "view": "[VIEW]", "component": "[CMP]",
    "class": "[CLS]", "external": "[EXT]",
}


def render_tree(node: FlowNode, prefix="", is_last=True, is_root=True):
    lines = []
    if is_root:
        lines.append("%s %s%s" % (TAG.get(node.kind, "[?]"), node.label,
                                  ("  " + node.note) if node.note else ""))
        child_prefix = ""
    else:
        branch = "`-- " if is_last else "|-- "
        lines.append("%s%s%s %s%s" % (prefix, branch, TAG.get(node.kind, "[?]"),
                                      node.label, ("  " + node.note) if node.note else ""))
        child_prefix = prefix + ("    " if is_last else "|   ")
    for i, c in enumerate(node.children):
        lines += render_tree(c, child_prefix, i == len(node.children) - 1, False)
    return lines


def _mid(s, counter):
    base = re.sub(r"[^0-9A-Za-z_]", "_", s)
    return "n%d_%s" % (counter, base[:40])


def render_mermaid(roots):
    out = ["```mermaid", "flowchart TD"]
    seen = {}
    counter = [0]
    edges = []

    def node_id(label, kind):
        if label in seen:
            return seen[label]
        counter[0] += 1
        nid = _mid(label, counter[0])
        seen[label] = nid
        shape = {"url": ('([\"', '\"])'), "sql": ('[(\"', '\")]'),
                 "view": ('>\"', '\"]')}.get(kind, ('[\"', '\"]'))
        out.append("    %s%s%s%s" % (nid, shape[0], label.replace('"', "'"), shape[1]))
        return nid

    def walk(n):
        nid = node_id(n.label, n.kind)
        for c in n.children:
            cid = walk(c)
            e = (nid, cid)
            if e not in edges:
                edges.append(e)
                out.append("    %s --> %s" % (nid, cid))
        return nid

    for r in roots:
        walk(r)
    out.append("```")
    return "\n".join(out)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def cmd_stats(idx, args):
    print("Java types        : %d  (scanned %d .java files)" % (len(idx.types), idx.n_java))
    print("  controllers     : %d" % sum(1 for t in idx.types.values() if t.is_controller()))
    print("  services        : %d" % sum(1 for t in idx.types.values() if t.is_service()))
    print("  repositories    : %d" % sum(1 for t in idx.types.values() if t.is_repository()))
    print("Mapper namespaces : %d  (scanned %d .xml files)" % (len(idx.mapper_ns), idx.n_xml))
    print("Mapped statements : %d" % len(idx.mapper_stmts))
    print("URL mappings      : %d" % len(idx.url_map))
    print("JSP pages         : %d" % idx.n_jsp)
    print("Parse errors      : %d" % len(idx.errors))
    if args.errors and idx.errors:
        print("\n-- parse errors --")
        for p, e in idx.errors[:50]:
            print("  %s\n    %s" % (p, e))


def cmd_urls(idx, args):
    rows = sorted(idx.url_map, key=lambda e: e[0])
    g = args.grep.lower() if args.grep else None
    for url, https, cfqn, mname in rows:
        line = "%-8s %-40s -> %s.%s" % ("/".join(https) or "ANY", url, cfqn, mname)
        if g and g not in line.lower():
            continue
        print(line)


def cmd_flow(idx, args):
    roots = []
    if args.url:
        roots, matches = trace_url(idx, args.url)
        if not matches:
            print("No controller mapping matched URL: %s" % args.url)
            print("Try `urls` to list available mappings, or use --method.")
            return
    elif args.method:
        fqn, _, mname = args.method.rpartition(".")
        if not fqn:
            print("Use fully-qualified --method, e.g. com.x.OrderController.save")
            return
        if fqn not in idx.types:
            cands = [f for f in idx.types if f.endswith("." + fqn) or f.split(".")[-1] == fqn]
            if len(cands) == 1:
                fqn = cands[0]
            else:
                print("Class not found: %s%s" % (fqn,
                      ("  candidates: " + ", ".join(cands)) if cands else ""))
                return
        roots = [trace_method(idx, fqn, mname, set())]
    elif args.jsp:
        page = idx.jsp.get(args.jsp) or idx.jsp.get(os.path.abspath(args.jsp))
        if page is None:
            cands = [p for p in idx.jsp if p.replace("\\", "/").endswith(args.jsp.replace("\\", "/"))]
            if len(cands) == 1:
                page = idx.jsp[cands[0]]
            else:
                print("JSP not found: %s%s" % (args.jsp,
                      ("  candidates: " + ", ".join(cands[:10])) if cands else ""))
                return
        print("JSP: %s" % page.path)
        if page.includes:
            print("  includes: %s" % ", ".join(page.includes))
        any_hit = False
        for u in page.urls:
            nu = idx.normalize_jsp_url(u, args.context_path)
            if not nu:
                continue
            r, matches = trace_url(idx, nu)
            if matches:
                any_hit = True
                roots += r
            else:
                print("  (no controller) %s" % u)
        if not any_hit and not roots:
            print("No JSP-referenced URL matched a controller mapping.")
            return
    else:
        print("Specify one of: --url / --method / --jsp")
        return

    if args.mermaid:
        print(render_mermaid(roots))
    else:
        for r in roots:
            print("\n".join(render_tree(r)))
            print()


def main(argv=None):
    ap = argparse.ArgumentParser(prog="jflow", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add_common(p):
        p.add_argument("--src", action="append", required=True,
                       help="source root (repeatable)")
        p.add_argument("-v", "--verbose", action="store_true")

    p = sub.add_parser("stats", help="index summary")
    add_common(p)
    p.add_argument("--errors", action="store_true", help="list parse errors")

    p = sub.add_parser("urls", help="list URL -> controller method mappings")
    add_common(p)
    p.add_argument("--grep", help="filter substring")
    p.add_argument("--context-path", default="")

    p = sub.add_parser("flow", help="trace a process flow")
    add_common(p)
    p.add_argument("--url", help="entry URL, e.g. /order/save")
    p.add_argument("--method", help="entry method FQN, e.g. com.x.OrderController.save")
    p.add_argument("--jsp", help="entry JSP path (suffix match ok)")
    p.add_argument("--context-path", default="", help="strip from JSP URLs before matching")
    p.add_argument("--mermaid", action="store_true", help="emit a Mermaid diagram")

    args = ap.parse_args(argv)
    for r in args.src:
        if not os.path.isdir(r):
            sys.stderr.write("ERROR: not a directory: %s\n" % r)
            return 2
    idx = build_index(args.src, verbose=args.verbose)
    {"stats": cmd_stats, "urls": cmd_urls, "flow": cmd_flow}[args.cmd](idx, args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
