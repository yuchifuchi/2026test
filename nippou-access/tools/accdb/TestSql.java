// 生成した .accdb に、web/include/sql.asp の SQL を実際に流して確かめる。
//
//   java TestSql <accdb> <sqls.json>
//
// UCanAccess (JDBC) で実行する。ACE そのものではないが、
// 「列名が違う」「別名が解決できない」「かっこが合わない」の類はここで落ちる。
import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.sql.*;
import java.util.*;

public class TestSql {
    public static void main(String[] args) throws Exception {
        String json = new String(Files.readAllBytes(new File(args[1]).toPath()),
                                 StandardCharsets.UTF_8);
        // { "名前": "SQL", ... } だけの単純な形を読む
        List<String[]> cases = new ArrayList<>();
        int i = json.indexOf('{') + 1;
        while (true) {
            int k1 = json.indexOf('"', i);
            if (k1 < 0) break;
            int k2 = endOfString(json, k1);
            String name = unescape(json.substring(k1 + 1, k2));
            int v1 = json.indexOf('"', json.indexOf(':', k2));
            int v2 = endOfString(json, v1);
            cases.add(new String[]{name, unescape(json.substring(v1 + 1, v2))});
            i = v2 + 1;
        }

        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        String url = "jdbc:ucanaccess://" + new File(args[0]).getAbsolutePath()
                   + ";openExclusive=false;ignoreCase=true";
        int ok = 0, ng = 0;
        try (Connection c = DriverManager.getConnection(url)) {
            for (String[] cs : cases) {
                String head = cs[1].trim().substring(0, 6).toUpperCase();
                if (!head.startsWith("SELECT")) {          // INSERT / UPDATE / DELETE
                    try (Statement st = c.createStatement()) {
                        int n = st.executeUpdate(cs[1]);
                        System.out.printf("  OK  %-22s %d 行に効いた%n", cs[0], n);
                        ok++;
                    } catch (Exception e) {
                        System.out.println("  NG  " + cs[0] + " : " + e.getMessage());
                        ng++;
                    }
                    continue;
                }
                try (Statement st = c.createStatement();
                     ResultSet rs = st.executeQuery(cs[1])) {
                    ResultSetMetaData m = rs.getMetaData();
                    int rows = 0;
                    StringBuilder first = new StringBuilder();
                    while (rs.next()) {
                        if (rows == 0) {
                            for (int j = 1; j <= Math.min(m.getColumnCount(), 8); j++) {
                                first.append(j > 1 ? ", " : "")
                                     .append(m.getColumnLabel(j)).append("=").append(rs.getString(j));
                            }
                        }
                        rows++;
                    }
                    System.out.printf("  OK  %-22s 列 %2d / 行 %3d%n",
                                      cs[0], m.getColumnCount(), rows);
                    if (first.length() > 0) System.out.println("        " + first);
                    ok++;
                } catch (Exception e) {
                    System.out.println("  NG  " + cs[0] + " : " + e.getMessage());
                    ng++;
                }
            }
        }
        System.out.println();
        System.out.println("通った " + ok + " / 落ちた " + ng);
        System.exit(ng == 0 ? 0 : 1);
    }

    static int endOfString(String s, int start) {
        int i = start + 1;
        while (true) {
            char c = s.charAt(i);
            if (c == '\\') { i += 2; continue; }
            if (c == '"') return i;
            i++;
        }
    }

    static String unescape(String s) {
        StringBuilder b = new StringBuilder();
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c != '\\') { b.append(c); continue; }
            char e = s.charAt(++i);
            switch (e) {
                case 'n': b.append('\n'); break;
                case 't': b.append('\t'); break;
                case 'r': b.append('\r'); break;
                case 'u': b.append((char) Integer.parseInt(s.substring(i + 1, i + 5), 16)); i += 4; break;
                default: b.append(e);
            }
        }
        return b.toString();
    }
}
