// db/schema.sql を読んで、Jackcess 4.0.5 で .accdb を作る。
//
//   java -cp "tools/lib/builder/*" tools/java/BuildAccdb.java <schema.sql> <出力.accdb> <検査用.json>
//
// なぜ自前で SQL を読むのか:
//   Windows も Access も無い環境で .accdb を作る手段が Jackcess しかなく、
//   Jackcess は SQL の DDL を実行できない。そこで「DDL と初期データを 1 本のスクリプトに書く」
//   という決まりを守るため、schema.sql の限られた書式だけを読む小さな読み取り器を持つ。
//
// 作ったあと、同じファイルを開き直して「実際に書かれた内容」を JSON に書き出す。
// その JSON を tools/check_accdb.py が schema.sql と独立に突き合わせる（読み取り器の誤りも見つかるように）。
import com.healthmarketscience.jackcess.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.util.*;

public class BuildAccdb {
    // ---------- 読み取ったモデル ----------
    static class Col { String name, type; int len; boolean notNull; }
    static class Idx { String name; List<String> cols = new ArrayList<>(); boolean pk, unique; }
    static class Fk { String name, table, refTable; List<String> cols = new ArrayList<>(), refCols = new ArrayList<>(); }
    static class Tbl { String name; List<Col> cols = new ArrayList<>(); List<Idx> idx = new ArrayList<>(); }
    static class Ins { String table; List<String> cols = new ArrayList<>(); List<Object> vals = new ArrayList<>(); }

    static Map<String, Tbl> tables = new LinkedHashMap<>();
    static List<Fk> fks = new ArrayList<>();
    static List<Ins> inserts = new ArrayList<>();

    // ---------- 字句 ----------
    static List<String> tokens(String s) {
        List<String> t = new ArrayList<>();
        int i = 0, n = s.length();
        while (i < n) {
            char c = s.charAt(i);
            if (Character.isWhitespace(c) || c == '﻿') { i++; continue; }
            if (c == '-' && i + 1 < n && s.charAt(i + 1) == '-') { while (i < n && s.charAt(i) != '\n') i++; continue; }
            if (c == '[') { int j = s.indexOf(']', i); if (j < 0) fail("] がありません"); t.add(s.substring(i, j + 1)); i = j + 1; continue; }
            if (c == '\'') {
                StringBuilder b = new StringBuilder("'"); i++;
                while (true) {
                    if (i >= n) fail("文字列が閉じていません");
                    char d = s.charAt(i);
                    if (d == '\'') { if (i + 1 < n && s.charAt(i + 1) == '\'') { b.append('\''); i += 2; continue; } i++; break; }
                    b.append(d); i++;
                }
                t.add(b.toString()); continue;
            }
            if (c == '#') { int j = s.indexOf('#', i + 1); t.add(s.substring(i, j + 1)); i = j + 1; continue; }
            if ("(),;".indexOf(c) >= 0) { t.add(String.valueOf(c)); i++; continue; }
            if (Character.isLetterOrDigit(c) || c == '-' || c == '_') {
                int j = i + 1;
                while (j < n && (Character.isLetterOrDigit(s.charAt(j)) || s.charAt(j) == '_' || s.charAt(j) == '.')) j++;
                t.add(s.substring(i, j)); i = j; continue;
            }
            fail("読めない文字: " + c);
        }
        return t;
    }

    static List<String> tk; static int p;
    static String peek() { return p < tk.size() ? tk.get(p) : ""; }
    static String next() { if (p >= tk.size()) fail("文が途中で終わっています"); return tk.get(p++); }
    static boolean isKw(String s, String kw) { return s.equalsIgnoreCase(kw); }
    static void expect(String kw) { String s = next(); if (!isKw(s, kw)) fail("「" + kw + "」が必要な所に「" + s + "」があります（" + p + " 語目）"); }
    static String name() { String s = next(); if (!s.startsWith("[")) fail("名前は [ ] で囲んでください: " + s); return s.substring(1, s.length() - 1); }
    static List<String> nameList() {
        expect("("); List<String> r = new ArrayList<>();
        while (true) { r.add(name()); if (isKw(peek(), "ASC") || isKw(peek(), "DESC")) next(); String s = next(); if (s.equals(")")) break; if (!s.equals(",")) fail(", か ) が必要です"); }
        return r;
    }
    static void fail(String m) { throw new RuntimeException("schema.sql: " + m); }

    static void parse(String src) {
        tk = tokens(src); p = 0;
        while (p < tk.size()) {
            String s = next();
            if (s.equals(";")) continue;
            if (isKw(s, "CREATE")) {
                String k = next();
                if (isKw(k, "TABLE")) parseTable();
                else if (isKw(k, "INDEX")) parseIndex(false);
                else if (isKw(k, "UNIQUE")) { expect("INDEX"); parseIndex(true); }
                else fail("CREATE の後ろが読めません: " + k);
            } else if (isKw(s, "ALTER")) {
                expect("TABLE"); String t = name(); expect("ADD"); expect("CONSTRAINT");
                Fk f = new Fk(); f.name = name(); f.table = t; expect("FOREIGN"); expect("KEY");
                f.cols = nameList(); expect("REFERENCES"); f.refTable = name(); f.refCols = nameList(); fks.add(f);
            } else if (isKw(s, "INSERT")) {
                expect("INTO"); Ins in = new Ins(); in.table = name(); in.cols = nameList(); expect("VALUES"); expect("(");
                while (true) {
                    String v = next();
                    if (v.startsWith("'")) in.vals.add(v.substring(1));
                    else if (isKw(v, "NULL")) in.vals.add(null);
                    else if (isKw(v, "TRUE")) in.vals.add(Boolean.TRUE);
                    else if (isKw(v, "FALSE")) in.vals.add(Boolean.FALSE);
                    else if (v.startsWith("#")) in.vals.add(LocalDateTime.parse(v.substring(1, v.length() - 1) + "T00:00:00"));
                    else in.vals.add(Integer.valueOf(v));
                    String d = next(); if (d.equals(")")) break; if (!d.equals(",")) fail("値の区切りが読めません");
                }
                if (in.cols.size() != in.vals.size()) fail(in.table + " への INSERT で列と値の数が合いません");
                inserts.add(in);
            } else fail("文の始まりが読めません: " + s);
            if (p < tk.size() && !peek().equals(";")) fail("文の終わりに ; がありません（" + peek() + "）");
        }
    }

    static void parseTable() {
        Tbl t = new Tbl(); t.name = name(); expect("(");
        if (tables.containsKey(t.name)) fail("表が 2 回定義されています: " + t.name);
        while (true) {
            String s = peek();
            if (isKw(s, "CONSTRAINT")) {
                next(); String cn = name(); String k = next();
                if (isKw(k, "UNIQUE")) { Idx x = new Idx(); x.name = cn; x.unique = true; x.cols = nameList(); t.idx.add(x); }
                else if (isKw(k, "PRIMARY")) { expect("KEY"); Idx x = new Idx(); x.name = cn; x.pk = true; x.unique = true; x.cols = nameList(); t.idx.add(x); }
                else if (isKw(k, "FOREIGN")) { expect("KEY"); Fk f = new Fk(); f.name = cn; f.table = t.name; f.cols = nameList(); expect("REFERENCES"); f.refTable = name(); f.refCols = nameList(); fks.add(f); }
                else fail("CONSTRAINT の種類が読めません: " + k);
            } else {
                Col c = new Col(); c.name = name(); String ty = next().toUpperCase(Locale.ROOT);
                c.type = ty;
                if (ty.equals("TEXT")) { expect("("); c.len = Integer.parseInt(next()); expect(")"); }
                else if (!Arrays.asList("LONG", "COUNTER", "MEMO", "BIT", "DATETIME").contains(ty)) fail("使えない型です: " + ty);
                while (true) {
                    String q = peek();
                    if (isKw(q, "NOT")) { next(); expect("NULL"); c.notNull = true; }
                    else if (isKw(q, "CONSTRAINT")) {
                        next(); String cn = name(); expect("PRIMARY"); expect("KEY");
                        Idx x = new Idx(); x.name = cn; x.pk = true; x.unique = true; x.cols.add(c.name); t.idx.add(x);
                    } else break;
                }
                t.cols.add(c);
            }
            String d = next(); if (d.equals(")")) break; if (!d.equals(",")) fail(t.name + " の定義の区切りが読めません: " + d);
        }
        tables.put(t.name, t);
    }

    static void parseIndex(boolean unique) {
        Idx x = new Idx(); x.name = name(); x.unique = unique; expect("ON"); String t = name(); x.cols = nameList();
        Tbl tb = tables.get(t); if (tb == null) fail("索引の対象の表が先に定義されていません: " + t);
        tb.idx.add(x);
    }

    // ---------- 書き込み ----------
    static DataType dt(Col c) {
        switch (c.type) {
            case "LONG": case "COUNTER": return DataType.LONG;
            case "TEXT": return DataType.TEXT;
            case "MEMO": return DataType.MEMO;
            case "BIT": return DataType.BOOLEAN;
            case "DATETIME": return DataType.SHORT_DATE_TIME;
        }
        throw new IllegalStateException(c.type);
    }

    static void build(File out) throws IOException {
        if (out.exists() && !out.delete()) throw new IOException("古いファイルを消せません: " + out);
        try (Database db = new DatabaseBuilder(out).setFileFormat(Database.FileFormat.V2007).create()) {
            for (Tbl t : tables.values()) {
                TableBuilder tb = new TableBuilder(t.name);
                for (Col c : t.cols) {
                    ColumnBuilder cb = new ColumnBuilder(c.name, dt(c));
                    if (c.type.equals("TEXT")) cb.setLengthInUnits(c.len);
                    if (c.type.equals("COUNTER")) cb.setAutoNumber(true);
                    // NOT NULL は Access の「値要求」で表す（Yes/No 型と オートナンバー型は常に値があるので付けない）
                    if (c.notNull && !c.type.equals("BIT") && !c.type.equals("COUNTER")) cb.putProperty(PropertyMap.REQUIRED_PROP, true);
                    // 空文字を許さない。画面側は「空なら Null」で保存するので、空文字が入るのは作りの誤りだけ。
                    if (c.type.equals("TEXT") || c.type.equals("MEMO")) cb.putProperty(PropertyMap.ALLOW_ZERO_LEN_PROP, false);
                    tb.addColumn(cb);
                }
                for (Idx x : t.idx) {
                    IndexBuilder ib = new IndexBuilder(x.name).addColumns(x.cols.toArray(new String[0]));
                    if (x.pk) ib.setPrimaryKey(); else if (x.unique) ib.setUnique();
                    tb.addIndex(ib);
                }
                tb.toTable(db);
            }
            for (Fk f : fks) {
                // Jackcess の RelationshipBuilder は (主＝「1」側, 従＝「多」側) の順で渡す
                RelationshipBuilder rb = new RelationshipBuilder(f.refTable, f.table).setName(f.name).setReferentialIntegrity();
                for (int i = 0; i < f.cols.size(); i++) rb.addColumns(f.refCols.get(i), f.cols.get(i));
                rb.toRelationship(db);
            }
            for (Ins in : inserts) {
                Table t = db.getTable(in.table);
                Map<String, Object> row = new LinkedHashMap<>();
                for (int i = 0; i < in.cols.size(); i++) row.put(in.cols.get(i), in.vals.get(i));
                t.addRowFromMap(row);
            }
        }
    }

    // ---------- 読み戻し（検査用） ----------
    static String js(String s) {
        StringBuilder b = new StringBuilder("\"");
        for (char c : s.toCharArray()) {
            if (c == '"' || c == '\\') b.append('\\').append(c);
            else if (c < 0x20) b.append(String.format("\\u%04x", (int) c));
            else b.append(c);
        }
        return b.append('"').toString();
    }

    static void dump(File f, File json) throws IOException {
        StringBuilder o = new StringBuilder();
        byte[] head = new byte[0x20];
        try (InputStream in = new FileInputStream(f)) { if (in.read(head) != head.length) throw new IOException("短すぎるファイル"); }
        try (Database db = new DatabaseBuilder(f).setReadOnly(true).open()) {
            o.append("{\"formatByte\":").append(head[0x14] & 0xff);
            o.append(",\"fileFormat\":").append(js(db.getFileFormat().name()));
            o.append(",\"tables\":[");
            boolean ft = true;
            for (String tn : db.getTableNames()) {
                Table t = db.getTable(tn);
                if (!ft) o.append(','); ft = false;
                o.append("{\"name\":").append(js(tn)).append(",\"rows\":").append(t.getRowCount()).append(",\"columns\":[");
                boolean fc = true;
                for (Column c : t.getColumns()) {
                    if (!fc) o.append(','); fc = false;
                    Object req = c.getProperties().getValue(PropertyMap.REQUIRED_PROP);
                    Object azl = c.getProperties().getValue(PropertyMap.ALLOW_ZERO_LEN_PROP);
                    o.append("{\"name\":").append(js(c.getName()))
                     .append(",\"type\":").append(js(c.getType().name()))
                     .append(",\"lengthInUnits\":").append(c.getLengthInUnits())
                     .append(",\"autoNumber\":").append(c.isAutoNumber())
                     .append(",\"required\":").append(Boolean.TRUE.equals(req))
                     .append(",\"allowZeroLength\":").append(azl == null ? "null" : azl.toString())
                     .append('}');
                }
                o.append("],\"indexes\":[");
                boolean fi = true;
                for (Index x : t.getIndexes()) {
                    if (!fi) o.append(','); fi = false;
                    o.append("{\"name\":").append(js(x.getName())).append(",\"pk\":").append(x.isPrimaryKey())
                     .append(",\"unique\":").append(x.isUnique()).append(",\"foreignKey\":").append(x.isForeignKey()).append(",\"columns\":[");
                    boolean fx = true;
                    for (Index.Column ic : x.getColumns()) { if (!fx) o.append(','); fx = false; o.append(js(ic.getName())); }
                    o.append("]}");
                }
                o.append("],\"data\":[");
                boolean fr = true;
                for (Row r : t) {
                    if (!fr) o.append(','); fr = false;
                    o.append('{'); boolean fv = true;
                    for (Map.Entry<String, Object> e : r.entrySet()) {
                        if (!fv) o.append(','); fv = false;
                        Object v = e.getValue();
                        o.append(js(e.getKey())).append(':');
                        if (v == null) o.append("null");
                        else if (v instanceof Boolean || v instanceof Number) o.append(v);
                        else o.append(js(v.toString()));
                    }
                    o.append('}');
                }
                o.append("]}");
            }
            o.append("],\"relationships\":[");
            boolean frl = true;
            for (Relationship r : db.getRelationships()) {
                if (!frl) o.append(','); frl = false;
                o.append("{\"name\":").append(js(r.getName()))
                 .append(",\"primaryTable\":").append(js(r.getFromTable().getName()))
                 .append(",\"foreignTable\":").append(js(r.getToTable().getName()))
                 .append(",\"integrity\":").append(r.hasReferentialIntegrity())
                 .append(",\"primaryColumns\":[");
                boolean fx = true;
                for (Column c : r.getFromColumns()) { if (!fx) o.append(','); fx = false; o.append(js(c.getName())); }
                o.append("],\"foreignColumns\":[");
                fx = true;
                for (Column c : r.getToColumns()) { if (!fx) o.append(','); fx = false; o.append(js(c.getName())); }
                o.append("]}");
            }
            o.append("]}");
        }
        Files.write(json.toPath(), o.toString().getBytes(StandardCharsets.UTF_8));
    }

    public static void main(String[] a) throws Exception {
        if (a.length != 3) { System.err.println("使い方: BuildAccdb <schema.sql> <出力.accdb> <検査用.json>"); System.exit(2); }
        parse(new String(Files.readAllBytes(Paths.get(a[0])), StandardCharsets.UTF_8));
        File out = new File(a[1]);
        build(out);
        dump(out, new File(a[2]));
        System.out.println("作成: " + out + "（表 " + tables.size() + "・関連 " + fks.size() + "・初期行 " + inserts.size() + "）");
    }
}
