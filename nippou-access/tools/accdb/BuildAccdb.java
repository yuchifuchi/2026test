// -*- coding: utf-8 -*-
// 電話応対日報 集計システム ― Access ファイル (.accdb) を作る
//
//   java BuildAccdb <schema.json> <出力先.accdb>
//
// tools/gen_accdb.py が src/*.bas から取り出した定義 (schema.json) を読んで、
// Jackcess で本物の .accdb (V2016) を書き出す。Windows も Access も要らない。
//
// 作るもの : テーブル・主キー・一意制約・インデックス・リレーションシップ・マスタ
// 作れないもの : 保存クエリ (Q_...)。Jackcess が対応していないため。
//                Web 版は保存クエリを使わない作りにしてある (web/include/sql.asp)。
import com.healthmarketscience.jackcess.*;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.regex.*;

public class BuildAccdb {

    // --- ごく小さな JSON 読み取り ------------------------------------------
    // 依存を増やしたくないので、この用途にだけ足りるものを自前で持つ。
    static class J {
        final String s;
        int i;
        J(String s) { this.s = s; }

        Object parse() { ws(); return value(); }
        void ws() { while (i < s.length() && Character.isWhitespace(s.charAt(i))) i++; }

        Object value() {
            char c = s.charAt(i);
            if (c == '{') return obj();
            if (c == '[') return arr();
            if (c == '"') return str();
            if (s.startsWith("true", i)) { i += 4; return Boolean.TRUE; }
            if (s.startsWith("false", i)) { i += 5; return Boolean.FALSE; }
            if (s.startsWith("null", i)) { i += 4; return null; }
            int st = i;
            while (i < s.length() && "-+.eE0123456789".indexOf(s.charAt(i)) >= 0) i++;
            String num = s.substring(st, i);
            return num.contains(".") ? (Object) Double.valueOf(num) : (Object) Long.valueOf(num);
        }

        Map<String, Object> obj() {
            Map<String, Object> m = new LinkedHashMap<>();
            i++; ws();
            if (s.charAt(i) == '}') { i++; return m; }
            while (true) {
                ws();
                String k = str();
                ws(); i++;                       // ':'
                ws();
                m.put(k, value());
                ws();
                if (s.charAt(i) == ',') { i++; continue; }
                i++;                             // '}'
                return m;
            }
        }

        List<Object> arr() {
            List<Object> l = new ArrayList<>();
            i++; ws();
            if (s.charAt(i) == ']') { i++; return l; }
            while (true) {
                ws();
                l.add(value());
                ws();
                if (s.charAt(i) == ',') { i++; continue; }
                i++;                             // ']'
                return l;
            }
        }

        String str() {
            StringBuilder b = new StringBuilder();
            i++;                                 // 開きの "
            while (true) {
                char c = s.charAt(i++);
                if (c == '"') return b.toString();
                if (c != '\\') { b.append(c); continue; }
                char e = s.charAt(i++);
                switch (e) {
                    case 'n': b.append('\n'); break;
                    case 't': b.append('\t'); break;
                    case 'r': b.append('\r'); break;
                    case 'b': b.append('\b'); break;
                    case 'f': b.append('\f'); break;
                    case 'u':
                        b.append((char) Integer.parseInt(s.substring(i, i + 4), 16));
                        i += 4;
                        break;
                    default: b.append(e);
                }
            }
        }
    }

    @SuppressWarnings("unchecked")
    static List<Object> list(Object o) { return (List<Object>) o; }
    @SuppressWarnings("unchecked")
    static Map<String, Object> map(Object o) { return (Map<String, Object>) o; }

    static DataType type(String t) {
        switch (t) {
            case "LONG":
            case "COUNTER": return DataType.LONG;
            case "TEXT":    return DataType.TEXT;
            case "MEMO":    return DataType.MEMO;
            case "DATETIME":return DataType.SHORT_DATE_TIME;
            case "BIT":     return DataType.BOOLEAN;
        }
        throw new IllegalArgumentException("知らない型: " + t);
    }

    static final Pattern DATE_ONLY = Pattern.compile("\\d{4}-\\d{2}-\\d{2}");

    static Object convert(Object v, String type) {
        if (v == null) return null;
        switch (type) {
            case "LONG":
            case "COUNTER":
                return (v instanceof Boolean) ? ((Boolean) v ? 1 : 0)
                                              : Integer.valueOf(((Number) v).intValue());
            case "BIT":
                return (v instanceof Boolean) ? v : (((Number) v).intValue() != 0);
            case "DATETIME": {
                String s = String.valueOf(v);
                return DATE_ONLY.matcher(s).matches()
                        ? LocalDate.parse(s).atStartOfDay()
                        : LocalDateTime.parse(s.replace(' ', 'T'));
            }
            default:
                return String.valueOf(v);
        }
    }

    public static void main(String[] args) throws IOException {
        if (args.length < 2) {
            System.err.println("使い方: BuildAccdb <schema.json> <出力先.accdb>");
            System.exit(2);
        }
        String json = new String(Files.readAllBytes(new File(args[0]).toPath()),
                                StandardCharsets.UTF_8);
        Map<String, Object> schema = map(new J(json).parse());
        File out = new File(args[1]);
        if (out.exists() && !out.delete()) {
            System.err.println("既存ファイルを消せません: " + out);
            System.exit(1);
        }
        out.getParentFile().mkdirs();

        // 形式は V2007 (Access 2007〜2016 と同じ「既定の .accdb」)。
        // V2016 で作ると、少し古い ACE では
        //   「このデータベースを開くには、Microsoft Access のより新しい
        //     バージョンが必要です」
        // になって開けない。新しい機能は何も使っていないので、
        // いちばん広く読める形にしておく。
        String fmt = System.getProperty("accdb.format", "V2007");
        Database db = DatabaseBuilder.create(
                Database.FileFormat.valueOf(fmt), out);

        // --- テーブル ---
        Map<String, List<Map<String, Object>>> colsOf = new LinkedHashMap<>();
        for (Object o : list(schema.get("tables"))) {
            Map<String, Object> t = map(o);
            String name = (String) t.get("name");
            TableBuilder tb = new TableBuilder(name);
            List<Map<String, Object>> cols = new ArrayList<>();
            for (Object co : list(t.get("columns"))) {
                Map<String, Object> c = map(co);
                cols.add(c);
                String ct = (String) c.get("type");
                ColumnBuilder cb = new ColumnBuilder((String) c.get("name"), type(ct));
                if ("COUNTER".equals(ct)) cb.setAutoNumber(true);
                if (c.get("len") != null) cb.setLengthInUnits(((Number) c.get("len")).intValue());
                tb.addColumn(cb);
            }
            colsOf.put(name, cols);

            List<Object> pk = list(t.get("pk"));
            if (!pk.isEmpty()) tb.setPrimaryKey(pk.toArray(new String[0]));
            for (Object uo : list(t.get("unique"))) {
                Map<String, Object> u = map(uo);
                tb.addIndex(new IndexBuilder((String) u.get("name"))
                        .addColumns(list(u.get("cols")).toArray(new String[0]))
                        .setUnique());
            }
            for (Object io : list(t.get("indexes"))) {
                Map<String, Object> ix = map(io);
                tb.addIndex(new IndexBuilder((String) ix.get("name"))
                        .addColumns(list(ix.get("cols")).toArray(new String[0])));
            }
            Table tbl = tb.toTable(db);

            // NOT NULL は Access の「値要求」として持たせる (VBA 版と同じ状態にする)
            for (Map<String, Object> c : cols) {
                if (!Boolean.TRUE.equals(c.get("notNull"))) continue;
                if (pk.contains(c.get("name"))) continue;      // 主キーは既に必須
                try {
                    PropertyMap p = tbl.getColumn((String) c.get("name")).getProperties();
                    p.put(PropertyMap.REQUIRED_PROP, true);
                    p.save();
                } catch (Exception e) {
                    System.out.println("  ! 値要求を設定できません: "
                            + name + "." + c.get("name") + " (" + e + ")");
                }
            }
            System.out.println("表: " + name + " (" + cols.size() + " 列)");
        }

        // --- リレーションシップ ---
        for (Object o : list(schema.get("relationships"))) {
            Map<String, Object> r = map(o);
            String parent = (String) r.get("parent");
            String child = (String) r.get("child");
            try {
                new RelationshipBuilder(parent, child)
                        .addColumns((String) r.get("parentColumn"), (String) r.get("childColumn"))
                        .setReferentialIntegrity()
                        .setName((String) r.get("name"))
                        .toRelationship(db);
                System.out.println("参照整合性: " + r.get("name")
                        + "  " + parent + " → " + child);
            } catch (Exception e) {
                System.out.println("  ! リレーションを作れません: " + r.get("name") + " (" + e + ")");
                throw e;
            }
        }

        // --- マスタの中身 ---
        Map<String, Object> data = map(schema.get("data"));
        for (Map.Entry<String, Object> e : data.entrySet()) {
            Table t = db.getTable(e.getKey());
            List<Map<String, Object>> cols = colsOf.get(e.getKey());
            List<Object[]> rows = new ArrayList<>();
            for (Object ro : list(e.getValue())) {
                Map<String, Object> row = map(ro);
                Object[] vals = new Object[cols.size()];
                for (int i = 0; i < cols.size(); i++) {
                    Map<String, Object> c = cols.get(i);
                    String cn = (String) c.get("name");
                    if ("COUNTER".equals(c.get("type")) && !row.containsKey(cn)) {
                        vals[i] = Column.AUTO_NUMBER;
                    } else {
                        vals[i] = convert(row.get(cn), (String) c.get("type"));
                    }
                }
                rows.add(vals);
            }
            t.addRows(rows);
            System.out.println("投入: " + e.getKey() + " " + rows.size() + " 行");
        }
        db.close();

        // --- 読み直して確かめる ---
        Database v = DatabaseBuilder.open(out);
        System.out.println();
        System.out.println("--- 確認 (書いたファイルを読み直しています) ---");
        System.out.println("形式: " + v.getFileFormat());
        for (String n : new TreeSet<>(v.getTableNames())) {
            Table t = v.getTable(n);
            StringBuilder idx = new StringBuilder();
            for (Index ix : t.getIndexes()) {
                idx.append(idx.length() > 0 ? ", " : "").append(ix.getName());
                if (ix.isPrimaryKey()) idx.append("(主キー)");
                else if (ix.isUnique()) idx.append("(一意)");
            }
            System.out.printf("  %-12s %2d 列 / %4d 行 / 索引: %s%n",
                    n, t.getColumnCount(), t.getRowCount(), idx);
        }
        System.out.println("リレーションシップ: " + v.getRelationships().size() + " 本");
        Table ops = v.getTable("M_担当者");
        int shown = 0;
        for (Row r : ops) {
            System.out.println("  例) " + r.get("担当者コード") + " " + r.get("氏名")
                    + " / " + r.get("職員区分"));
            if (++shown == 3) break;
        }
        v.close();
        System.out.println("OK");
    }
}
