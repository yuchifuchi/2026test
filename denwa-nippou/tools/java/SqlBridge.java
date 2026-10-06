// UCanAccess 5.0.1 で .accdb に SQL を流す「中継役」。Python の検査プログラムから使う。
//
//   java -cp "tools/lib/ucanaccess/*" tools/java/SqlBridge.java
//
// 標準入力から 1 行 1 件の JSON を受け取り、1 行 1 件の JSON を返す。
//   {"op":"open","path":"..."}                         接続
//   {"op":"query","sql":"...","params":[{"t":"int","v":1}, ...]}   SELECT
//   {"op":"exec", "sql":"...","params":[...]}           INSERT / UPDATE / DELETE
//   {"op":"begin"} {"op":"commit"} {"op":"rollback"} {"op":"close"}
// 値の型 t: int / dbl / str / date（"yyyy-MM-ddTHH:mm:ss"）/ bool / null
//
// 注意: UCanAccess で通る SQL が ACE（本番の Access）で通るとは限らない。
//       ACE だけが受け付けない書き方は tools/vbsim/sqlcheck.py で別に止めている。
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.sql.*;
import java.time.LocalDateTime;
import java.util.*;

public class SqlBridge {
    // ---------- 最小限の JSON ----------
    static int ip; static String is;
    static Object parse(String s) { is = s; ip = 0; Object v = val(); return v; }
    static void ws() { while (ip < is.length() && Character.isWhitespace(is.charAt(ip))) ip++; }
    static Object val() {
        ws(); char c = is.charAt(ip);
        if (c == '{') {
            ip++; Map<String, Object> m = new LinkedHashMap<>(); ws();
            if (is.charAt(ip) == '}') { ip++; return m; }
            while (true) { ws(); String k = (String) val(); ws(); ip++; /* : */ m.put(k, val()); ws(); char d = is.charAt(ip++); if (d == '}') return m; }
        }
        if (c == '[') {
            ip++; List<Object> l = new ArrayList<>(); ws();
            if (is.charAt(ip) == ']') { ip++; return l; }
            while (true) { l.add(val()); ws(); char d = is.charAt(ip++); if (d == ']') return l; }
        }
        if (c == '"') {
            ip++; StringBuilder b = new StringBuilder();
            while (true) {
                char d = is.charAt(ip++);
                if (d == '"') return b.toString();
                if (d == '\\') {
                    char e = is.charAt(ip++);
                    switch (e) {
                        case 'n': b.append('\n'); break; case 'r': b.append('\r'); break; case 't': b.append('\t'); break;
                        case 'b': b.append('\b'); break; case 'f': b.append('\f'); break;
                        case 'u': b.append((char) Integer.parseInt(is.substring(ip, ip + 4), 16)); ip += 4; break;
                        default: b.append(e);
                    }
                } else b.append(d);
            }
        }
        if (is.startsWith("true", ip)) { ip += 4; return Boolean.TRUE; }
        if (is.startsWith("false", ip)) { ip += 5; return Boolean.FALSE; }
        if (is.startsWith("null", ip)) { ip += 4; return null; }
        int j = ip; while (j < is.length() && "+-0123456789.eE".indexOf(is.charAt(j)) >= 0) j++;
        String n = is.substring(ip, j); ip = j;
        if (n.contains(".") || n.contains("e") || n.contains("E")) return Double.valueOf(n);
        return Long.valueOf(n);
    }
    static String js(Object v) {
        if (v == null) return "null";
        if (v instanceof Boolean || v instanceof Integer || v instanceof Long) return v.toString();
        if (v instanceof Number) { double d = ((Number) v).doubleValue(); return Double.isFinite(d) ? Double.toString(d) : "null"; }
        String s = v.toString();
        StringBuilder b = new StringBuilder("\"");
        for (char c : s.toCharArray()) {
            if (c == '"' || c == '\\') b.append('\\').append(c);
            else if (c < 0x20) b.append(String.format("\\u%04x", (int) c));
            else b.append(c);
        }
        return b.append('"').toString();
    }

    static Connection conn;

    static void bind(PreparedStatement ps, List<Object> params) throws SQLException {
        if (params == null) return;
        int i = 1;
        for (Object o : params) {
            @SuppressWarnings("unchecked") Map<String, Object> m = (Map<String, Object>) o;
            String t = (String) m.get("t"); Object v = m.get("v");
            switch (t) {
                case "null": ps.setNull(i, Types.VARCHAR); break;
                case "int": ps.setInt(i, ((Number) v).intValue()); break;
                case "dbl": ps.setDouble(i, ((Number) v).doubleValue()); break;
                case "str": ps.setString(i, (String) v); break;
                case "bool": ps.setBoolean(i, (Boolean) v); break;
                case "date": ps.setTimestamp(i, Timestamp.valueOf(LocalDateTime.parse((String) v))); break;
                default: throw new SQLException("知らない型: " + t);
            }
            i++;
        }
    }

    static String colType(int t) {
        switch (t) {
            case Types.INTEGER: case Types.SMALLINT: case Types.TINYINT: return "int";
            case Types.BIGINT: return "bigint";
            case Types.DOUBLE: case Types.FLOAT: case Types.REAL: case Types.DECIMAL: case Types.NUMERIC: return "dbl";
            case Types.BOOLEAN: case Types.BIT: return "bool";
            case Types.TIMESTAMP: case Types.DATE: case Types.TIME: return "date";
            default: return "str";
        }
    }

    static String handle(Map<String, Object> req) throws Exception {
        String op = (String) req.get("op");
        switch (op) {
            case "open": {
                Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
                // memory=true: 検査は小さなファイルで行うので全部メモリに載せる
                conn = DriverManager.getConnection("jdbc:ucanaccess://" + req.get("path") + ";memory=true;immediatelyReleaseResources=true");
                conn.setAutoCommit(true);
                return "{\"ok\":true}";
            }
            case "close": if (conn != null) conn.close(); conn = null; return "{\"ok\":true}";
            case "begin": conn.setAutoCommit(false); return "{\"ok\":true}";
            case "commit": conn.commit(); conn.setAutoCommit(true); return "{\"ok\":true}";
            case "rollback": conn.rollback(); conn.setAutoCommit(true); return "{\"ok\":true}";
            case "query": {
                @SuppressWarnings("unchecked") List<Object> params = (List<Object>) req.get("params");
                try (PreparedStatement ps = conn.prepareStatement((String) req.get("sql"))) {
                    bind(ps, params);
                    try (ResultSet rs = ps.executeQuery()) {
                        ResultSetMetaData md = rs.getMetaData();
                        int n = md.getColumnCount();
                        StringBuilder o = new StringBuilder("{\"ok\":true,\"columns\":[");
                        String[] types = new String[n];
                        for (int i = 1; i <= n; i++) {
                            types[i - 1] = colType(md.getColumnType(i));
                            if (i > 1) o.append(',');
                            o.append("{\"name\":").append(js(md.getColumnLabel(i))).append(",\"t\":").append(js(types[i - 1])).append('}');
                        }
                        o.append("],\"rows\":[");
                        boolean first = true;
                        while (rs.next()) {
                            if (!first) o.append(','); first = false;
                            o.append('[');
                            for (int i = 1; i <= n; i++) {
                                if (i > 1) o.append(',');
                                Object v = rs.getObject(i);
                                if (v == null) o.append("null");
                                else if (v instanceof Timestamp) o.append(js(((Timestamp) v).toLocalDateTime().toString()));
                                else if (v instanceof java.sql.Date) o.append(js(((java.sql.Date) v).toLocalDate().atStartOfDay().toString()));
                                else if (v instanceof java.math.BigDecimal) o.append(js(((java.math.BigDecimal) v).doubleValue()));
                                else o.append(js(v));
                            }
                            o.append(']');
                        }
                        return o.append("]}").toString();
                    }
                }
            }
            case "exec": {
                @SuppressWarnings("unchecked") List<Object> params = (List<Object>) req.get("params");
                try (PreparedStatement ps = conn.prepareStatement((String) req.get("sql"))) {
                    bind(ps, params);
                    int n = ps.executeUpdate();
                    return "{\"ok\":true,\"affected\":" + n + "}";
                }
            }
        }
        throw new Exception("知らない指示: " + op);
    }

    public static void main(String[] a) throws Exception {
        BufferedReader in = new BufferedReader(new InputStreamReader(System.in, StandardCharsets.UTF_8));
        PrintStream out = new PrintStream(new FileOutputStream(FileDescriptor.out), true, "UTF-8");
        String line;
        while ((line = in.readLine()) != null) {
            if (line.trim().isEmpty()) continue;
            String res;
            try {
                @SuppressWarnings("unchecked") Map<String, Object> req = (Map<String, Object>) parse(line);
                res = handle(req);
            } catch (Throwable e) {
                Throwable c = e; StringBuilder m = new StringBuilder();
                while (c != null) { if (m.length() > 0) m.append(" / "); m.append(c.getClass().getSimpleName()).append(": ").append(c.getMessage()); c = c.getCause(); }
                res = "{\"ok\":false,\"error\":" + js(m.toString()) + "}";
            }
            out.println(res);
        }
    }
}
