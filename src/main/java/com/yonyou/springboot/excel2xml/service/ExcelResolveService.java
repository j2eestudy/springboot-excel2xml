package com.yonyou.springboot.excel2xml.service;

import com.yonyou.springboot.excel2xml.utils.FileCache;
import com.yonyou.springboot.excel2xml.vo.ShowVO;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.output.Format;
import org.jdom.output.XMLOutputter;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Author: shijq
 * @Date: 2019/3/6 18:42
 */
@Service
public class ExcelResolveService {

    public static Map<String, Cell> mergedRegionMap = null;

    public static ThreadLocal<Integer> colLocal = new ThreadLocal<>();
    public static ThreadLocal<Integer> rowLocal = new ThreadLocal<>();

    public static ThreadLocal<String> cellFormulaLocal = new ThreadLocal<>();

    public FormulaEvaluator evaluator = null;

    @Value("${file.downloadPath}")
    private String downloadPath;

    private String path;

    public List<ShowVO> readExcell(InputStream stream, String path) {
        List<ShowVO> showVOS = new ArrayList<>();
        try {
            XSSFWorkbook wb = new XSSFWorkbook(stream);
            evaluator = wb.getCreationHelper().createFormulaEvaluator();
            //获取工作薄的个数，即一个excel文件中包含了多少个Sheet工作簿
            int WbLength = wb.getNumberOfSheets();
            //对每一个工作簿进行操作
            for (int i = 0; i < WbLength; i++) {
                mergedRegionMap = new HashMap();
                XSSFSheet shee = wb.getSheetAt(i);
                String filePath = path + "\\" + shee.getSheetName() + ".xml";
                File file = new File(filePath);

                String uuid = UUID.randomUUID().toString();

                ShowVO vo = new ShowVO();
                vo.setFileName(file.getName());
                vo.setUrl(downloadPath+uuid);
                vo.setUuid(uuid);

                FileCache.set(uuid,filePath);

                showVOS.add(vo);

                FileOutputStream fo = new FileOutputStream(file);// 得到输入流
                Element root = new Element("cells");
                Document doc = new Document(root);

                int length = shee.getLastRowNum();

                for (int j = 0; j < length; j++) {
                    XSSFRow row = shee.getRow(j);
                    if (row == null) {
                        continue;
                    }
                    int cellNum = row.getPhysicalNumberOfCells();// 获取一行中最后一个单元格的位置

                    //确定最后一列位置
                    if (j == 0) {
                        Integer lastCol = getLastColCell(row, cellNum);
                        if (lastCol != null) {
                            cellNum = lastCol;
                        }
                    }

                    //判断是否最后一行
                    boolean isLastRow = checkThisRowIsBreak(row);
                    if (isLastRow) {
                        break;
                    }


                    Element rowElement = new Element("row");
                    root.addContent(rowElement);
                    for (int k = 0; k < cellNum; k++) {
                        XSSFCell cell = row.getCell((short) k);

                        if (cell == null) {
                            cellNum++;//如果存在空列，那么cellNum增加1，这一步很重要。
                            continue;
                        } else {

                            String mergedRegion = isMergedRegion(shee, j, k);
                            if (mergedRegion == null) {
                                Element element = setValueElements(cell);
                                rowElement.addContent(element);
                            } else {
                                if (mergedRegionMap.containsKey(mergedRegion)) {
                                    continue;
                                } else {
                                    mergedRegionMap.put(mergedRegion, cell);
                                    String[] strs = mergedRegion.split(",");
                                    colLocal.set(Integer.valueOf(strs[1]) - Integer.valueOf(strs[0]));
                                    rowLocal.set(Integer.valueOf(strs[3]) - Integer.valueOf(strs[2]));
                                    Element element = setValueElements(cell);
                                    rowElement.addContent(element);
                                }
                            }


                        }

                    }

                }

                Format format = Format.getRawFormat().setEncoding("UTF-8").setIndent(" ");
                XMLOutputter XMLOut = new XMLOutputter(format);// 在元素后换行，每一层元素缩排四格
                XMLOut.output(doc, fo);
                fo.close();

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return showVOS;
    }

    private Element setValueElements(XSSFCell cell) {

        Element element = new Element("cell");
        List<Element> elementList = new ArrayList<Element>();


        /**
         * 备注：value必须在fml之前
         */
        String[] elements = {"cellname", "formid", "enabled", "validate", "quarter", "keyword", "datatype",
                "format", "canimport", "condition", "isCombox", "colspan", "rowspan", "color", "value",
                "fml", "isshow", "width", "istitle", "ismeasure", "iseidter", "validatemsg"};
        List<String> names = Arrays.asList(elements);
        for (int i = 0; i < names.size(); i++) {
            Element item = new Element(names.get(i));
            if ("value".equals(names.get(i))) {
                setValue(item, cell);
            } else if ("formid".equals(names.get(i))) {
                item.setText(cell.getAddress().toString());
            } else if ("isshow".equals(names.get(i))) {
                item.setText("true");
            } else if ("colspan".equals(names.get(i))) {
                if (colLocal.get() != null) {
                    item.setText(String.valueOf(colLocal.get() + 1));
                    colLocal.remove();
                } else {
                    item.setText("1");
                }
            } else if ("rowspan".equals(names.get(i))) {
                if (rowLocal.get() != null) {
                    item.setText(String.valueOf(rowLocal.get() + 1));
                    rowLocal.remove();
                } else {
                    item.setText("1");
                }
            } else if ("fml".equals(names.get(i))) {
                if (cellFormulaLocal.get() != null) {
                    item.setText(cellFormulaLocal.get());
                    cellFormulaLocal.remove();
                } else {
                    item.setText("");
                }
            } else {
                item.setText("");
            }
            elementList.add(item);
        }
        element.addContent(elementList);
        return element;

    }

    private void setValue(Element item, XSSFCell cell) {
        String cellvalue = "";
        switch (cell.getCellType()) {
            // 如果当前Cell的Type为NUMERIC
            case HSSFCell.CELL_TYPE_NUMERIC:
                double value = cell.getNumericCellValue();
                item.setText(String.valueOf(value));

            case HSSFCell.CELL_TYPE_FORMULA: {
                //System.out.println(cell.getCellFormula());
                // 判断当前的cell是否为Date
                try {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);
                        item.setText(cellvalue);
                    } else {
                        cellvalue = String.valueOf(cell.getCellFormula());
                        cellFormulaLocal.set(cellvalue);
                        double dd = evaluator.evaluate(cell).getNumberValue();
                        item.setText(String.valueOf(dd));
                        System.out.println(cellvalue);
                    }
                } catch (Exception e) {
                    try {
                        cellvalue = String.valueOf(cell.getCellFormula());
                        double dd = evaluator.evaluate(cell).getNumberValue();
                        item.setText(String.valueOf(dd));
                        System.out.println(cellvalue);
                        cellFormulaLocal.set(cellvalue);
                    } catch (Exception e1) {
                        try {
                            cellvalue = String.valueOf(cell.getNumericCellValue());
                            item.setText(cellvalue);
                        } catch (Exception e2) {
                            cellvalue = String.valueOf(cell.getStringCellValue());
                            item.setText(cellvalue);
                        }
                    }

                }

                break;
            }
            // 如果当前Cell的Type为STRIN
            case HSSFCell.CELL_TYPE_STRING:
                // 取得当前的Cell字符串
                cellvalue = cell.getRichStringCellValue()
                        .getString();
                item.setText(cellvalue);
                break;
            // 默认的Cell值
            default:
                cellvalue = "";
                item.setText(cellvalue);

        }

    }

    /**
     * 判断是否最后一行
     *
     * @param row
     * @return
     */
    private Boolean checkThisRowIsBreak(XSSFRow row) {
        boolean isTrue = false;
        XSSFCell cell = row.getCell(0);
        String cellValue = getCellValue(cell);
        if ("END".equals(cellValue)) {
            isTrue = true;
        }
        return isTrue;
    }


    /**
     * 线下人工excel处理，第一行的最后一列加END内容，获取对应的index
     *
     * @param row
     * @param cellNum
     * @return
     */
    public Integer getLastColCell(XSSFRow row, int cellNum) {
        Integer endCol = null;
        for (int k = 0; k < cellNum; k++) {
            XSSFCell cell = row.getCell((short) k);
            String cellValue = getCellValue(cell);
            if ("END".equals(cellValue)) {
                endCol = k;
            }
        }
        return endCol;
    }

    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public String getCellValue(Cell cell) {

        if (cell == null) return "";

        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

            return cell.getStringCellValue();

        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

            return String.valueOf(cell.getBooleanCellValue());

        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

            return cell.getCellFormula();

        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param sheet
     * @param row    行下标
     * @param column 列下标
     * @return
     */
    private String isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    String key = firstColumn + "," + lastColumn + "," + firstRow + "," + lastRow;
                    return key;
                }
            }
        }
        return null;
    }

}
