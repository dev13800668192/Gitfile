package until;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.sql.*;
import java.util.*;

/**
 * @author by
 * @version 1.0
 * @date 2020/1/16 13:44
 */
public class PoiUntil {

    private final static Logger LOGGER = LoggerFactory.getLogger(PoiUntil.class);
    private static  String DRIVER ;
    private static  String URL ;
    private static  String USERNAME;
    private static  String PASSWORD ;
    private static String databaseName ;


    private static final String SQL = "SELECT * FROM ";
    static {
        Properties properties = new Properties();
        String configFile = "jdbc.properties";
        try{
            InputStream is = PoiUntil.class.getClassLoader().getResourceAsStream(configFile);
            BufferedReader bf =new BufferedReader(new InputStreamReader(is,"gbk"));
            properties.load(bf);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        DRIVER = properties.getProperty("jdbc.driver");
        URL = properties.getProperty("jdbc.url");
        USERNAME = properties.getProperty("jdbc.username");
        PASSWORD = properties.getProperty("jdbc.password");
        databaseName=properties.getProperty("jdbc.databaseName");

        try {
            Class.forName(DRIVER);
        } catch (ClassNotFoundException e) {
            LOGGER.error("can not load jdbc driver", e);
        }
    }



    public  String getDatabaseName() {
        return databaseName;
    }


    /**
     * 获取数据库连接
     *
     * @return
     */
    public  Connection getConnection() {
        Connection conn = null;
        try {
            conn = DriverManager.getConnection(URL,USERNAME, PASSWORD);
        } catch (SQLException e) {
           LOGGER.error("连接数据库失败", e);
            e.printStackTrace();
        }
        return conn;
    }

    /**
     * 关闭数据库连接
     * @param conn
     */
    public  void closeConnection(Connection conn) {
        if(conn != null) {
            try {
                conn.close();
            } catch (SQLException e) {
                LOGGER.error("关闭数据库失败", e);
            }
        }
    }

    /**
     * 获取数据库下的所有表名
     */
    public  List<String> getTableNames() {
        List<String> tableNames = new ArrayList<String>();
        Connection conn = getConnection();
        ResultSet rs = null;
        try {
            DatabaseMetaData db = conn.getMetaData();
            rs = db.getTables(null, null, null, new String[] { "TABLE" });
            while(rs.next()) {
                tableNames.add(rs.getString(3));
            }
        } catch (SQLException e) {
            LOGGER.error("获取数据库下所有表名失败", e);
        } finally {
            try {
                rs.close();
                closeConnection(conn);
            } catch (SQLException e) {
                LOGGER.error("关闭数据库失败", e);
            }
        }
        return tableNames;
    }

    /**
     * 获取表中所有字段名称
     * @param tableName 表名
     * @return
     */
    public  List<String> getColumnNames(String tableName) {
        List<String> columnNames = new ArrayList<String>();
        Connection conn = getConnection();
        PreparedStatement pStemt = null;
        String tableSql = SQL + tableName;
        try {
            pStemt = conn.prepareStatement(tableSql);
            ResultSetMetaData rsmd = pStemt.getMetaData();
            int size = rsmd.getColumnCount();
            for (int i = 0; i < size; i++) {
                columnNames.add(rsmd.getColumnName(i + 1));
            }
        } catch (SQLException e) {
            LOGGER.error("获取表中所有字段名称失败", e);
        } finally {
            if (pStemt != null) {
                try {
                    pStemt.close();
                    closeConnection(conn);
                } catch (SQLException e) {
                    LOGGER.error("获取列名关闭 pstem 和连接失败", e);
                }
            }
        }
        return columnNames;
    }

    /**
     * 获取表中所有字段类型
     * @param tableName
     * @return
     */
    public  List<String> getColumnTypes(String tableName) {
        List<String> columnTypes = new ArrayList<String>();
        Connection conn = getConnection();
        PreparedStatement pStemt = null;
        String tableSql = SQL + tableName;
        try {
            pStemt = conn.prepareStatement(tableSql);
            ResultSetMetaData rsmd = pStemt.getMetaData();
            int size = rsmd.getColumnCount();
            for (int i = 0; i < size; i++) {
                columnTypes.add(rsmd.getColumnTypeName(i + 1));
            }
        } catch (SQLException e) {
            LOGGER.error("获取表中所有字段类型", e);
        } finally {
            if (pStemt != null) {
                try {
                    pStemt.close();
                    closeConnection(conn);
                } catch (SQLException e) {
                    LOGGER.error("获取字段类型关闭 pstem 和连接失败", e);
                }
            }
        }
        return columnTypes;
    }

    /**
     * 获取表中所有内容
     * @param tableName
     */
    public  List<Map<String,Object>> getColumnDate(String tableName) {
        List<Map<String ,Object>> columnData = new ArrayList<Map<String, Object>>();
        Connection conn = getConnection();
        PreparedStatement pStemt = null;
        String tableSql = SQL + tableName;
        try {
            pStemt = conn.prepareStatement(tableSql);
            ResultSet rs = pStemt.executeQuery();
            ResultSetMetaData rsmd = pStemt.getMetaData();
            int size = rsmd.getColumnCount();
            while(rs.next()){
                Map<String,Object> map =new LinkedHashMap<String, Object>();
                for (int i=1; i<=size;i++){
                    map.put(rsmd.getColumnName(i),rs.getObject(i));
                }
                columnData.add(map);
            }

        } catch (SQLException e) {
            LOGGER.error("获取数据库数据失败", e);
        } finally {
            if (pStemt != null) {
                try {
                    pStemt.close();
                    closeConnection(conn);
                } catch (SQLException e) {
                    LOGGER.error("关闭数据库失败", e);
                }
            }
        }
        return columnData;
    }

    /**
     * 获取表中表结构
     * @param tableName
     */

    public  List<Map<String,Object>> getColumnStruct(String tableName) {
        List<Map<String,Object>> columnStruct = new ArrayList<Map<String,Object>>();
        Connection conn = getConnection();
        ResultSet rs = null;
        try {
            DatabaseMetaData db = conn.getMetaData();
            rs = db.getColumns(null, "%", tableName, "%");
            while(rs.next()) {
                Map<String,Object> map =new LinkedHashMap<String, Object>();
                String columnName =rs.getString("COLUMN_NAME");
                String columnType =rs.getString("TYPE_NAME");
                String dataSize  =rs.getString("COLUMN_SIZE");
                String nullable =rs.getString("NULLABLE");
                String def=rs.getString("COLUMN_DEF");
                String remarks =rs.getString("REMARKS");
                String TypeSize =columnType.toLowerCase()+"("+dataSize+")";
                String key =getColumnKey(tableName);
                map.put("columnName",columnName);
                map.put("columnType",TypeSize);
                if (nullable.equals(1)){
                    map.put("nullable","YES");
                }else {
                    map.put("nullable","NO");
                }
                map.put("def",def);
                if (key!=null&&key.equals(columnName)){
                    map.put("key","Y");
                }else {
                    map.put("key","");
                }
                map.put("remrk",remarks);
                columnStruct.add(map);
            }
        } catch (SQLException e) {
            LOGGER.error("获取表结构失败", e);
        } finally {
            try {
                rs.close();
                closeConnection(conn);
            } catch (SQLException e) {
                LOGGER.error("关闭连接失败", e);
            }
        }
        return columnStruct;
    }

    /**
     * 获取表中表中主键
     * @param tableName
     */

    public  String getColumnKey(String tableName) {
        String PrimaryKey =null;
        List<Map<String,Object>> columnKey = new ArrayList<Map<String,Object>>();
        Connection conn = getConnection();
        ResultSet rs = null;
        try {
            DatabaseMetaData db = conn.getMetaData();
            rs = db.getPrimaryKeys(null, "%", tableName);
            if (rs!=null){
                while(rs.next()) {
                  PrimaryKey =rs.getString("COLUMN_NAME");
                }
            }
        } catch (SQLException e) {
            LOGGER.error("获取表中表中主键失败", e);
        } finally {
            try {
                rs.close();
                closeConnection(conn);
            } catch (SQLException e) {
                LOGGER.error("关闭连接失败", e);
            }
        }
        return PrimaryKey;
    }


    /**
     * 设置表头样式及参数
     * @param
     */
    public HSSFCellStyle getHeaderStyle(HSSFWorkbook wb){
        HSSFCellStyle headStyle=wb.createCellStyle();
        headStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //边框设置
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setBorderBottom(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);

        //字体设置
        HSSFFont font =wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)12);
        headStyle.setFont(font);
        return headStyle;
    }

    /**
     * 设置普通样式及参数
     * @param
     */
    public HSSFCellStyle getStyle(HSSFWorkbook wb){
        HSSFCellStyle style=wb.createCellStyle();
        //边框设置
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        //字体设置
        HSSFFont font =wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)12);
        style.setFont(font);
        return style;
    }

    /**
     * 设置超连接样式
     * @param
     */
    public HSSFCellStyle getHyperStyle(HSSFWorkbook wb){
        HSSFCellStyle hyperStyel=wb.createCellStyle();
        //边框设置
        hyperStyel.setBorderTop(BorderStyle.THIN);
        hyperStyel.setBorderBottom(BorderStyle.THIN);
        hyperStyel.setBorderLeft(BorderStyle.THIN);
        hyperStyel.setBorderRight(BorderStyle.THIN);
        hyperStyel.setAlignment(HorizontalAlignment.CENTER);

        //字体设置
        HSSFFont font =wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)12);
        hyperStyel.setFont(font);
        return hyperStyel;
    }
}
