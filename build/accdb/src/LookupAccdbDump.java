import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

public class LookupAccdbDump {
    private static final int DEFAULT_MAX_TABLES = 40;
    private static final int DEFAULT_MAX_ROWS = 200;
    private static final int DEFAULT_MAX_COLS = 40;
    private static final int MAX_CELL_CHARS = 4000;

    private static final class TableDump {
        String name;
        List<String> columns = new ArrayList<>();
        List<List<String>> rows = new ArrayList<>();
    }

    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            System.err.println("usage: LookupAccdbDump <sourcePath> <outJsonPath> [maxTables] [maxRows] [maxCols]");
            System.exit(2);
            return;
        }

        String sourcePath = args[0];
        String outJsonPath = args[1];
        int maxTables = parsePositiveInt(args, 2, DEFAULT_MAX_TABLES);
        int maxRows = parsePositiveInt(args, 3, DEFAULT_MAX_ROWS);
        int maxCols = parsePositiveInt(args, 4, DEFAULT_MAX_COLS);

        List<TableDump> dumps = readDatabase(sourcePath, maxTables, maxRows, maxCols);
        writeJson(outJsonPath, dumps);
    }

    private static int parsePositiveInt(String[] args, int index, int fallback) {
        if (index >= args.length) {
            return fallback;
        }
        try {
            int value = Integer.parseInt(args[index]);
            return value > 0 ? value : fallback;
        } catch (NumberFormatException ex) {
            return fallback;
        }
    }

    private static List<TableDump> readDatabase(String sourcePath, int maxTables, int maxRows, int maxCols) throws SQLException {
        List<TableDump> dumps = new ArrayList<>();
        String url = "jdbc:ucanaccess://" + sourcePath + ";memory=false;immediatelyReleaseResources=true";
        try (Connection conn = DriverManager.getConnection(url)) {
            DatabaseMetaData meta = conn.getMetaData();
            List<String> tableNames = new ArrayList<>();
            try (ResultSet tables = meta.getTables(null, null, "%", new String[]{"TABLE"})) {
                while (tables.next() && tableNames.size() < maxTables) {
                    String tableName = safeTrim(tables.getString("TABLE_NAME"));
                    if (!tableName.isEmpty()) {
                        tableNames.add(tableName);
                    }
                }
            }

            for (String tableName : tableNames) {
                TableDump dump = readTable(conn, tableName, maxRows, maxCols);
                if (dump != null && !dump.columns.isEmpty() && !dump.rows.isEmpty()) {
                    dumps.add(dump);
                }
            }
        }
        return dumps;
    }

    private static TableDump readTable(Connection conn, String tableName, int maxRows, int maxCols) {
        String escapedName = tableName.replace("]", "]]");
        String sql = "SELECT * FROM [" + escapedName + "]";
        TableDump dump = new TableDump();
        dump.name = tableName;
        try (Statement st = conn.createStatement(); ResultSet rs = st.executeQuery(sql)) {
            ResultSetMetaData meta = rs.getMetaData();
            int colCount = Math.min(maxCols, Math.max(0, meta.getColumnCount()));
            for (int i = 1; i <= colCount; i++) {
                dump.columns.add(safeTrim(meta.getColumnLabel(i)));
            }
            while (rs.next() && dump.rows.size() < maxRows) {
                List<String> row = new ArrayList<>();
                for (int i = 1; i <= colCount; i++) {
                    Object value = rs.getObject(i);
                    String text = value == null ? "" : String.valueOf(value);
                    if (text.length() > MAX_CELL_CHARS) {
                        text = text.substring(0, MAX_CELL_CHARS);
                    }
                    row.add(text);
                }
                dump.rows.add(row);
            }
            return dump;
        } catch (SQLException ex) {
            return null;
        }
    }

    private static void writeJson(String outJsonPath, List<TableDump> dumps) throws IOException {
        Path output = Path.of(outJsonPath);
        if (output.getParent() != null) {
            Files.createDirectories(output.getParent());
        }
        try (BufferedWriter writer = Files.newBufferedWriter(output, StandardCharsets.UTF_8)) {
            writer.write("{\"tables\":[");
            for (int i = 0; i < dumps.size(); i++) {
                TableDump table = dumps.get(i);
                if (i > 0) {
                    writer.write(",");
                }
                writer.write("{\"name\":\"");
                writer.write(jsonEscape(table.name));
                writer.write("\",\"columns\":[");
                for (int c = 0; c < table.columns.size(); c++) {
                    if (c > 0) {
                        writer.write(",");
                    }
                    writer.write("\"");
                    writer.write(jsonEscape(table.columns.get(c)));
                    writer.write("\"");
                }
                writer.write("],\"rows\":[");
                for (int r = 0; r < table.rows.size(); r++) {
                    if (r > 0) {
                        writer.write(",");
                    }
                    List<String> row = table.rows.get(r);
                    writer.write("[");
                    for (int c = 0; c < row.size(); c++) {
                        if (c > 0) {
                            writer.write(",");
                        }
                        writer.write("\"");
                        writer.write(jsonEscape(row.get(c)));
                        writer.write("\"");
                    }
                    writer.write("]");
                }
                writer.write("]}");
            }
            writer.write("]}");
        }
    }

    private static String safeTrim(String value) {
        return value == null ? "" : value.trim();
    }

    private static String jsonEscape(String value) {
        if (value == null || value.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder(value.length() + 16);
        for (int i = 0; i < value.length(); i++) {
            char ch = value.charAt(i);
            switch (ch) {
                case '\\':
                    sb.append("\\\\");
                    break;
                case '"':
                    sb.append("\\\"");
                    break;
                case '\b':
                    sb.append("\\b");
                    break;
                case '\f':
                    sb.append("\\f");
                    break;
                case '\n':
                    sb.append("\\n");
                    break;
                case '\r':
                    sb.append("\\r");
                    break;
                case '\t':
                    sb.append("\\t");
                    break;
                default:
                    if (ch < 0x20) {
                        sb.append(String.format("\\u%04x", (int) ch));
                    } else {
                        sb.append(ch);
                    }
                    break;
            }
        }
        return sb.toString();
    }
}
