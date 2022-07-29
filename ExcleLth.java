import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;

import java.io.File;
import java.util.ArrayList;

/**
 * @program:
 * @description: excel处理，编写程序对比不同 excel ，找到满足要求的内容。要求 1）以excel1为标准，在excel2中进行查询，找到相同的内容；2）以excel1为标准，在excel2中进行查询，对不同的内容输出打印。
 * @author: WZY
 * @create: 2022-06-28 23:27
 **/


public class ExcleLth {
    static int count = 0;
    // ?????12
    static ArrayList<String> col0Biaozhun = new ArrayList<>();
    static ArrayList<String> col1Biaozhun = new ArrayList<>();

    // ??????12
    static ArrayList<String> col0Daichazhao = new ArrayList<>();
    static ArrayList<String> col1Daichazhao = new ArrayList<>();
    public static void ReadExcelBiaozhun(String url) throws Exception {
        WorkbookSettings workbookSettings = new WorkbookSettings();
        workbookSettings.setEncoding("ISO-8859-1");
        Workbook workbook= Workbook.getWorkbook(new File(url),workbookSettings);
        
        // Workbook workbook = Workbook.getWorkbook(new File(url));
        Sheet sheet = workbook.getSheet(0);

        // 读取每一列，如果为""则说明结束
        for(int j = 1; j < sheet.getRows(); j++){ // 获得行数
            Cell cell = sheet.getCell(0, j); //  0:列  j:行
            if(cell.getContents().equals("")){
                break;
            }
            //System.out.println(cell.getContents() + " ");
            col0Biaozhun.add(cell.getContents());
        }
        // ?????????
        for(int j = 1; j < sheet.getRows(); j++){ // ?
            Cell cell = sheet.getCell(1, j); //  0:?  j:?
            if(cell.getContents().equals("")){
                break;
            }
            //System.out.println(cell.getContents() + " ");
            col1Biaozhun.add(cell.getContents());
        }
        workbook.close();
    }

    public static void ReadExcelDaichazhao(String url) throws Exception {


        Workbook workbook = Workbook.getWorkbook(new File(url));
        Sheet sheet = workbook.getSheet(0);

        // ?????????
        for(int j = 1; j < sheet.getRows(); j++){ // ?
            Cell cell = sheet.getCell(0, j); //  0:?  j:?
            if(cell.getContents().equals("")){
                break;
            }
            col0Daichazhao.add(cell.getContents());
        }

        // ?????????
        for(int j = 1; j < sheet.getRows(); j++){ // ?
            Cell cell = sheet.getCell(1, j); //  0:?  j:?
            if(cell.getContents().equals("")){
                break;
            }
            col1Daichazhao.add(cell.getContents());
        }
        workbook.close();
    }

    public static void main(String[] args) throws Exception {
        ReadExcelBiaozhun("C:\\Users\\Wzy\\Desktop\\LTH\\biaozhun.xls");
        ReadExcelDaichazhao("C:\\Users\\Wzy\\Desktop\\LTH\\daichazhao.xls");
        for(int i = 0; i < col0Biaozhun.size(); i++){
            String name = col0Biaozhun.get(i);
            for (int j = 0; j < col0Daichazhao.size(); j++){
                // ??????
                if(name.equals(col0Daichazhao.get(j))){
                    // ???
                    if(col1Biaozhun.get(i).equals(col1Daichazhao.get(j))){
                        int row = i + 2;
                        System.out.println("电缆名称 & 型号均相等 " + "| 当前手册中已查找到：" + "第 "+ row + " 行 | 电缆名称为："+  col0Biaozhun.get(i));
                        count++;
                    }else {
                        int row = i + 2;
                        System.out.println("电缆名称相等，但型号不等" + "| 当前手册中已查找到：" + "第 " + row + "行 | 电缆名称为：" + col0Biaozhun.get(i) + " | 出错电缆型号为：" + col1Daichazhao.get(j));
                        count++;
                    }

                }
            }
        }
        System.out.println("总共找到"+ count + "行数据");
    }
}
