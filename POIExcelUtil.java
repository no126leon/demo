import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Stack;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.mongodb.BasicDBObject;
import com.mongodb.util.JSON;
public class POIExcelUtil
{
private JSONObject dealExcel(String url, String outUrl, JSONObject jsonObject, InfoVo info, String bean, Long companyId, String token)
    {
        int coloumNum = 0;
        JSONArray picArr = new JSONArray();
        try
        {
            InputStream is = OSSUtil.getInstance().getFile(Constant.FLIE_LIBRARY_NAME, url);
            Workbook wb = WorkbookFactory.create(is);
            Sheet sheet = wb.getSheetAt(0);
            wb.setSheetName(0, " ");
            for (int i = 1; i < wb.getNumberOfSheets();)
            {
                wb.removeSheetAt(1);
            }
            int rownum = sheet.getFirstRowNum();
            coloumNum = sheet.getRow(rownum).getPhysicalNumberOfCells();
            List<CellRangeAddress> list = getCombineCellList(sheet);
            Iterator<Row> rowIterator = sheet.iterator();
            String localId = info.getLocalId();
            JSONObject enableFields = LayoutUtilAdapter.getEnableFields(bean, companyId.toString(), "0", localId);
            JSONArray subfieldArray = enableFields.getJSONArray("layout");
            JSONObject fieldJson = getFieldByLayout(subfieldArray);
            
            int maxRow = 0, aprsIndex = 0;
            List<Integer> blankRowIndex = new ArrayList<>();
            List<JSONObject> pictureJsons = new ArrayList<>();
            Map<String, List<JSONObject>> subfromMap = new LinkedHashMap<>();
            Map<String, JSONObject> approvalFieldMap = getApprovalFieldMap();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                if (row != null)
                {
                    String currentSubName = null;
                    JSONObject rowSubformDataJson = new JSONObject();
                    // 循环每一列
                    while (cellIterator.hasNext())
                    {
                        Cell cell = cellIterator.next();
                        int rowIndex = cell.getRowIndex();
                        for (int tr = (maxRow + 1); tr < rowIndex; tr++)
                        {
                            blankRowIndex.add(tr);
                        }
                        maxRow = rowIndex;
                        String value = getCellValue(cell);
                        JSONObject tempJson = isCombineCell(list, cell, sheet);
                        if (tempJson.getBooleanValue("flag"))
                        {
                            CellStyle style = cell.getCellStyle();
                            short borderLeft = tempJson.getShortValue("borderLeft");
                            short borderRight = tempJson.getShortValue("borderRight");
                            style.setBorderLeft(BorderStyle.valueOf(borderLeft));
                            style.setBorderRight(BorderStyle.valueOf(borderRight));
                            cell.setCellStyle(style);
                        }
                        if (StringUtils.isNotBlank(value))
                        {
                            // 正则表达式匹配 ${产品明细}*${产品名称}
                            String px = "\\$\\{(.*?)\\}";
                            int subFlag = value.indexOf('*');
                            Pattern PATTERN = Pattern.compile(px);
                            Matcher m = PATTERN.matcher(value);
                            List<String> vList = new ArrayList<>();
                            while (m.find())
                            {
                                vList.add(m.group(1));
                            }
                            // 如果只匹配到1个，说明是主表字段
                            if (subFlag < 0 && vList.size() > 0)
                            {
                                StringBuilder valueSB = new StringBuilder();
                                for (String fLabel : vList)
                                {
                                    String fName = fieldJson.getString(fLabel);
                                    if (fName == null)
                                    {
                                        continue;
                                    }
                                    if (fName.contains(TableConstant.TYPE_PICTURE) || fName.contains(TableConstant.TYPE_ATTACHMENT) || fName.contains(TableConstant.TYPE_SIGNATURE))
                                    {
                                        JSONArray pictures = jsonObject.getJSONArray(fName);
                                        for (int pi = 0; pi < pictures.size(); pi++)
                                        {
                                            JSONObject thePicJson = new JSONObject();
                                            thePicJson.put("picRow", (rowIndex + 1));
                                            thePicJson.put("picColumn", cell.getColumnIndex());
                                            thePicJson.put("picRatio", "");
                                            thePicJson.put("picSign", fName.contains(TableConstant.TYPE_SIGNATURE));
                                            thePicJson.put("picUrl", pictures.getJSONObject(pi).getString("file_url"));
                                            pictureJsons.add(thePicJson);
                                        }
                                    }
                                    else
                                    {
                                        valueSB.append(DataUtil.dealExportParam(fName, jsonObject.getString(fName)).toString());
                                    }
                                }
                                cell.setCellValue(valueSB.toString());
                            }
                            // 如果匹配到2个，说明是子表字段 第一个元素是子表单名称，第二个元素是子表字段名
                            else if (subFlag > 0 && vList.size() == 2)
                            {
                                cell.setCellValue("");
                                String subLabel = vList.get(0);
                                String fLabel = vList.get(1);
                                String subform = subLabel;
                                JSONObject cellJson = new JSONObject();
                                cellJson.put("down", value.contains("*u"));
                                cellJson.put("right", value.contains("*r"));
                                cellJson.put("index", cell.getColumnIndex());
                                // 流程审批
                                if (subLabel.equals(TableConstant.IF_TABLE_APPROVAL))
                                {
                                    JSONObject fjson = approvalFieldMap.get(fLabel);
                                    cellJson.put("field", fjson.getString("name"));
                                    cellJson.put("fieldType", fjson.getString("type"));
                                }
                                else
                                {
                                    subform = fieldJson.getString(subLabel);
                                    JSONObject subfields = getSubFieldByLayout(subfieldArray, subLabel);
                                    cellJson.put("field", subfields.getString(fLabel));
                                }
                                String subformName = rowSubformDataJson.getString("subform");
                                if (StringUtils.isBlank(subform) || (subformName != null && !subformName.equals(subform)))
                                {
                                    break;
                                }
                                currentSubName = subform;
                                cellJson.put("listCombineCell", isCombineCell(list, cell, sheet));
                                rowSubformDataJson.put("subform", subform);
                                rowSubformDataJson.put("rowIndex", rowIndex);
                                JSONArray cellArr = rowSubformDataJson.getJSONArray("cells");
                                if (cellArr == null)
                                {
                                    cellArr = new JSONArray();
                                    rowSubformDataJson.put("cells", cellArr);
                                }
                                cellArr.add(cellJson);
                            }
                        }
                        else
                        {
                            cell.setCellValue("");
                        }
                    }
                    if (currentSubName != null)
                    {
                        List<JSONObject> datas = subfromMap.get(currentSubName);
                        if (datas == null)
                        {
                            datas = new ArrayList<>();
                            subfromMap.put(currentSubName, datas);
                        }
                        datas.add(rowSubformDataJson);
                    }
                }
            }
            for (Integer rowIndex : blankRowIndex)
            {
                Row row = sheet.createRow(rowIndex);
                for (int i = 0; i < coloumNum; i++)
                {
                    row.createCell(i);
                }
            }
            int addRows = 0;
            List<Integer> keys = new ArrayList<>();
            for (Map.Entry<String, List<JSONObject>> entry : subfromMap.entrySet())
            {
                String subform = entry.getKey();
                List<JSONObject> meteDatas = entry.getValue();
                if (meteDatas == null || meteDatas.isEmpty())
                {
                    continue;
                }
                List<JSONObject> dataLS = null;
                if (subform.equals(TableConstant.IF_TABLE_APPROVAL))
                {
                    long dataId = jsonObject.getLongValue("id");
                    int dataType = jsonObject.getIntValue(TableConstant.FIELD_DATA_TYPE);
                    dataLS = getApprovalData(companyId, dataId, dataType, bean);
                }
                else
                {
                    JSONArray subdataArr = jsonObject.getJSONArray(subform);
                    if (subdataArr != null)
                    {
                        dataLS = subdataArr.toJavaList(JSONObject.class);
                    }
                }
                if (dataLS == null || dataLS.isEmpty())
                {
                    break;
                }
                int cellNums = 0;
                int added = addRows;
                int fsubCellIndex = 0;
                boolean right = false;
                int rowSindex = meteDatas.get(0).getIntValue("rowIndex");
                int rowEindex = meteDatas.get(meteDatas.size() - 1).getIntValue("rowIndex");
                int subrowSize = rowEindex - rowSindex + 1;
                for (int i = 0; i < dataLS.size(); i++)
                {
                    JSONObject subDataJson = dataLS.get(i);
                    if (i > 0 && !right)
                    {
                        addRows = addRows + subrowSize;
                    }
                    keys.add(aprsIndex + addRows);
                    for (JSONObject mete : meteDatas)
                    {
                        int rowIndex = mete.getIntValue("rowIndex");
                        JSONArray cellArr = mete.getJSONArray("cells");
                        Row row = sheet.getRow(rowIndex + addRows);
                        if (i == 0)
                        {
                            Iterator<Cell> cellIterator = row.cellIterator();
                            while (cellIterator.hasNext())
                            {
                                Cell cell = cellIterator.next();
                                cellNums = cell.getColumnIndex();
                            }
                        }
                        int csize = cellNums - fsubCellIndex + 1;
                        if (i > 0)
                        {
                            if (right)
                            {
                                for (int t = 1; t <= (cellNums - fsubCellIndex); t++)
                                {
                                    Cell cell = row.createCell((i - 1) * csize + t);
                                    Cell fromCell = row.getCell(fsubCellIndex + t - 1);
                                    cell.setCellStyle(fromCell.getCellStyle());
                                    cell.setCellValue(fromCell.getStringCellValue());
                                }
                            }
                            else
                            {
                                row = createRow(sheet, rowIndex + addRows);
                                Row oldrow = sheet.getRow(rowIndex + added);
                                Iterator<Cell> cellIterator = oldrow.cellIterator();
                                while (cellIterator.hasNext())
                                {
                                    Cell cell = cellIterator.next();
                                    Cell ncell = row.createCell(cell.getColumnIndex());
                                    ncell.setCellStyle(cell.getCellStyle());
                                    ncell.setCellValue(cell.getStringCellValue());
                                }
                            }
                        }
                        for (int j = 0; j < cellArr.size(); j++)
                        {
                            JSONObject cellJson = cellArr.getJSONObject(j);
                            int index = cellJson.getIntValue("index");
                            String field = cellJson.getString("field");
                            String fieldType = cellJson.getString("fieldType");
                            Cell cell = row.getCell(index);
                            if (i == 0)
                            {
                                right = cellJson.getBooleanValue("right");
                                if (j == 0)
                                {
                                    fsubCellIndex = index;
                                }
                            }
                            else if (right)
                            {
                                int nindex = (i - 1) * csize + (index - fsubCellIndex);
                                cell = row.getCell(nindex);
                            }
                            JSONObject listCombineCell = cellJson.getJSONObject("listCombineCell");
                            boolean flag = listCombineCell.getBooleanValue("flag");
                            String relValue = "";
                            String tempValue = subDataJson.getString(field);
                            if (StringUtils.isNotBlank(field))
                            {
                                if (fieldType != null && fieldType.equals(TableConstant.TYPE_PICTURE) && StringUtils.isNotBlank(tempValue))
                                {
                                    JSONObject thePicJson = new JSONObject();
                                    thePicJson.put("picRow", (rowIndex + 1));
                                    thePicJson.put("picColumn", cell.getColumnIndex());
                                    thePicJson.put("picRatio", "");
                                    thePicJson.put("picSign", true);
                                    thePicJson.put("picUrl", tempValue);
                                    pictureJsons.add(thePicJson);
                                }
                                else if (fieldType != null && fieldType.equals("approval_status"))
                                {
                                    relValue = AllEnum.ProcessApproveTaskEnum.getDesc4value(Util.parseToInteger(tempValue, 0));
                                }
                                else
                                {
                                    relValue = DataUtil.dealExportParam(fieldType, tempValue).toString();
                                }
                            }
                            cell.setCellValue(relValue);
                            if (flag)
                            {
                                int mergedRow = listCombineCell.getIntValue("mergedRow");
                                int mergedCol = listCombineCell.getIntValue("mergedCol");
                                short borderLeft = listCombineCell.getShortValue("borderLeft");
                                short borderRight = listCombineCell.getShortValue("borderRight");
                                CellRangeAddress cra =
                                    new CellRangeAddress(row.getRowNum(), row.getRowNum() + mergedRow - 1, cell.getColumnIndex(), cell.getColumnIndex() + mergedCol - 1);
                                sheet.addMergedRegion(cra);
                                RegionUtil.setBorderLeft(BorderStyle.valueOf(borderLeft), cra, sheet);
                                RegionUtil.setBorderRight(BorderStyle.valueOf(borderRight), cra, sheet);
                            }
                        }
                    }
                }
            }
            maxRow = maxRow + addRows;
            // 输出文件
            FileOutputStream fOut = new FileOutputStream(outUrl);
            wb.write(fOut);
            fOut.flush();
            // 操作结束，关闭文件
            fOut.close();
            wb.close();
            if (!keys.isEmpty())
            {
                Collections.sort(keys);
            }
            Map<String, PictureData> map = getPictures(sheet, list, keys);
            if (map != null && !map.isEmpty())
            {
                picArr = printImg(map, bean, companyId, maxRow, token);
            }
            if (!pictureJsons.isEmpty())
            {
                picArr.addAll(pictureJsons);
            }
        }
        catch (Exception e)
        {
            LOG.error(e.getMessage(), e);
        }
        JSONObject dataJson = new JSONObject();
        dataJson.put("url", outUrl);
        dataJson.put("picArr", picArr);
        dataJson.put("maxColumn", coloumNum);
        return dataJson;
    }
    
    /**
     * @param sheet
     * @return
     * @Description: 获取合并单元格集合
     */
    public List<CellRangeAddress> getCombineCellList(Sheet sheet)
    {
        List<CellRangeAddress> list = new ArrayList<>();
        // 获得一个 sheet 中合并单元格的数量
        int sheetmergerCount = sheet.getNumMergedRegions();
        // 遍历所有的合并单元格
        for (int i = 0; i < sheetmergerCount; i++)
        {
            // 获得合并单元格保存进list中
            CellRangeAddress ca = sheet.getMergedRegion(i);
            list.add(ca);
        }
        return list;
    }
    
    /**
     * 判断cell是否为合并单元格，是的话返回合并行数和列数（只要在合并区域中的cell就会返回合同行列数，但只有左上角第一个有数据）
     * 
     * @param listCombineCell 上面获取的合并区域列表
     * @param cell
     * @param sheet
     * @return
     * @throws Exception
     */
    public static JSONObject isCombineCell(List<CellRangeAddress> listCombineCell, Cell cell, Sheet sheet)
        throws Exception
    {
        int firstC = 0;
        int lastC = 0;
        int firstR = 0;
        int lastR = 0;
        int mergedRow = 0;
        int mergedCol = 0;
        JSONObject result = new JSONObject();
        result.put("flag", false);
        for (CellRangeAddress ca : listCombineCell)
        {
            // 获得合并单元格的起始行, 结束行, 起始列, 结束列
            firstC = ca.getFirstColumn();
            lastC = ca.getLastColumn();
            firstR = ca.getFirstRow();
            lastR = ca.getLastRow();
            // 判断cell是否在合并区域之内，在的话返回true和合并行列数
            if (cell.getRowIndex() >= firstR && cell.getRowIndex() <= lastR)
            {
                if (cell.getColumnIndex() >= firstC && cell.getColumnIndex() <= lastC)
                {
                    Row frow = sheet.getRow(firstR);
                    Cell lcell = frow.getCell(firstC);
                    Cell rcell = frow.getCell(lastC);
                    mergedRow = lastR - firstR + 1;
                    mergedCol = lastC - firstC + 1;
                    result.put("flag", true);
                    result.put("mergedRow", mergedRow);
                    result.put("mergedCol", mergedCol);
                    result.put("borderLeft", lcell.getCellStyle() == null ? 0 : lcell.getCellStyle().getBorderLeftEnum().getCode());
                    result.put("borderRight", rcell.getCellStyle() == null ? 0 : rcell.getCellStyle().getBorderRightEnum().getCode());
                    break;
                }
            }
        }
        return result;
    }
    
    private static JSONObject cellRatio(Sheet sheet, List<CellRangeAddress> listCombineCell, int row, int column)
        throws Exception
    {
        JSONObject result = null;
        for (CellRangeAddress ca : listCombineCell)
        {
            // 获得合并单元格的起始行, 结束行, 起始列, 结束列
            int firstC = ca.getFirstColumn();
            int lastC = ca.getLastColumn();
            int firstR = ca.getFirstRow();
            int lastR = ca.getLastRow();
            // 判断cell是否在合并区域之内，在的话返回true和合并行列数
            if (row >= firstR && row <= lastR)
            {
                if (column >= firstC && column <= lastC)
                {
                    result = new JSONObject();
                    float twidth = 0, cwidth = 0;
                    for (int i = firstC; i <= lastC; i++)
                    {
                        float ccw = sheet.getColumnWidthInPixels(i);
                        if (i < column)
                        {
                            cwidth += ccw;
                        }
                        twidth += ccw;
                    }
                    result.put("ratio", (cwidth / twidth));
                    result.put("firstC", firstC);
                    break;
                }
            }
        }
        return result;
    }
    
    /**
     * 获取图片和位置
     * 
     * @param sheet
     * @return
     * @throws IOException
     */
    public static Map<String, PictureData> getPictures(Sheet sheet, List<CellRangeAddress> listCombineCell, List<Integer> keys)
    {
        Map<String, PictureData> map = new HashMap<String, PictureData>();
        try
        {
            if (sheet instanceof HSSFSheet)
            {
                if (sheet.getDrawingPatriarch() != null)
                {
                    List<HSSFShape> list = ((HSSFSheet)sheet).getDrawingPatriarch().getChildren();
                    for (HSSFShape shape : list)
                    {
                        if (shape instanceof HSSFPicture)
                        {
                            HSSFPicture picture = (HSSFPicture)shape;
                            HSSFClientAnchor cAnchor = (HSSFClientAnchor)picture.getAnchor();
                            PictureData pdata = picture.getPictureData();
                            int row = cAnchor.getRow1();
                            int col = cAnchor.getCol1();
                            if (keys != null && keys.size() > 0)
                            {
                                if (row > keys.get(keys.size() - 1))
                                {
                                    row = row + keys.size();
                                }
                                else if (row >= keys.get(0))
                                {
                                    for (int i = 0; i < keys.size(); i++)
                                    {
                                        if (keys.get(i) >= row)
                                        {
                                            row = row + i + 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            String key = null;
                            JSONObject json = cellRatio(sheet, listCombineCell, row, col);
                            if (json != null)
                            {
                                key = (row + 1) + "_" + json.getString("firstC") + "_" + json.getString("ratio"); // 行号-列-比例
                                map.put(key, pdata);
                            }
                            else
                            {
                                key = (row + 1) + "_" + (col + 1); // 行号-列号
                            }
                            map.put(key, pdata);
                        }
                    }
                }
            }
            else
            {
                List<POIXMLDocumentPart> list = ((XSSFSheet)sheet).getRelations();
                for (POIXMLDocumentPart part : list)
                {
                    if (part instanceof XSSFDrawing)
                    {
                        XSSFDrawing drawing = (XSSFDrawing)part;
                        List<XSSFShape> shapes = drawing.getShapes();
                        for (XSSFShape shape : shapes)
                        {
                            XSSFPicture picture = (XSSFPicture)shape;
                            XSSFClientAnchor anchor = picture.getPreferredSize();
                            CTMarker marker = anchor.getFrom();
                            int row = marker.getRow();
                            int col = marker.getCol();
                            if (keys != null)
                            {
                                if (row > keys.get(keys.size() - 1))
                                {
                                    row = row + keys.size();
                                }
                                else if (row >= keys.get(0))
                                {
                                    for (int i = 0; i < keys.size(); i++)
                                    {
                                        if (keys.get(i) >= row)
                                        {
                                            row = row + i + 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            String key = null;
                            JSONObject json = cellRatio(sheet, listCombineCell, row, col);
                            if (json != null)
                            {
                                key = (row + 1) + "_" + json.getString("firstC") + "_" + json.getString("ratio"); // 行号-列-比例
                                
                            }
                            else
                            {
                                key = (row + 1) + "_" + (col + 1); // 行号-列号
                            }
                            map.put(key, picture.getPictureData());
                        }
                    }
                }
            }
        }
        catch (Exception e)
        {
            LOG.error(e.getMessage(), e);
        }
        return map;
    }
    
    /**
     * @param sheetList
     * @throws IOException
     * @Description:复制图片
     */
    public JSONArray printImg(Map<String, PictureData> sheetList, String bean, Long companyId, int maxRow, String token)
        throws IOException
    {
        JSONArray array = new JSONArray();
        Object key[] = sheetList.keySet().toArray();
        String url = Constant.PRINT_PREVIEW_URL;
        String prefix = url.concat(companyId.toString()).concat("_").concat(bean).concat("_");
        for (int i = 0; i < sheetList.size(); i++)
        {
            // 获取图片流
            PictureData pic = sheetList.get(key[i]);
            
            // 获取图片索引
            String picName = key[i].toString();
            String[] strArr = picName.split("_");
            
            // 获取图片格式
            String ext = pic.suggestFileExtension();
            byte[] data = pic.getData();
            
            // 图片保存路径
            String picUrl = prefix.concat(picName).concat(".").concat(ext);
            File imageFile = new File(picUrl);
            if (imageFile.exists())
            {
                imageFile.delete();
            }
            FileOutputStream out = new FileOutputStream(picUrl);
            out.write(data);
            out.close();
            
            int applyId = moduleAppService.getApplicationId(bean, token);
            String filePathUrl = "/common/file/download?";
            String fileUrlPath = FileDaoUtil.addApplyFile(bean, imageFile.getName(), imageFile.length(), String.valueOf(applyId), token, null, null);
            OSSUtil.getInstance().addFile(companyId, Constant.FLIE_LIBRARY_NAME, new FileInputStream(imageFile), fileUrlPath, imageFile.length());
            filePathUrl = filePathUrl.concat("bean=").concat(bean).concat("&fileName=").concat(fileUrlPath);
            
            JSONObject dataJson = new JSONObject();
            dataJson.put("picRow", Integer.valueOf(strArr[0]));
            dataJson.put("picColumn", Integer.valueOf(strArr[1]));
            dataJson.put("picRatio", strArr.length >= 3 ? strArr[2] : "");
            dataJson.put("picUrl", filePathUrl);
            dataJson.put("maxRow", Integer.valueOf(maxRow));
            array.add(dataJson);
        }
        return array;
    }
    
    /**
     * 根据excel单元格类型获取excel单元格值
     * 
     * @param cell
     * @return
     */
    @SuppressWarnings("deprecation")
    private String getCellValue(Cell cell)
    {
        String cellvalue = "";
        if (cell != null)
        {
            // 判断当前Cell的Type
            switch (cell.getCellType())
            {
                // 如果当前Cell的Type为NUMERIC
                case HSSFCell.CELL_TYPE_NUMERIC:
                {
                    short format = cell.getCellStyle().getDataFormat();
                    if (format == 14 || format == 31 || format == 57 || format == 58)
                    { // excel中的时间格式
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        double value = cell.getNumericCellValue();
                        Date date = DateUtil.getJavaDate(value);
                        cellvalue = sdf.format(date);
                    }
                    // 判断当前的cell是否为Date
                    else if (HSSFDateUtil.isCellDateFormatted(cell))
                    { // 先注释日期类型的转换，在实际测试中发现HSSFDateUtil.isCellDateFormatted(cell)只识别2014/02/02这种格式。
                      // 如果是Date类型则，取得该Cell的Date值 // 对2014-02-02格式识别不出是日期格式
                        Date date = cell.getDateCellValue();
                        DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = formater.format(date);
                    }
                    else
                    { // 如果是纯数字
                      // 取得当前Cell的数值
                        cellvalue = NumberToTextConverter.toText(cell.getNumericCellValue());
                        
                    }
                    break;
                }
                // 如果当前Cell的Type为STRIN
                case HSSFCell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getStringCellValue().replaceAll("'", "''");
                    break;
                case HSSFCell.CELL_TYPE_BLANK:
                    cellvalue = null;
                    break;
                // 默认的Cell值
                default:
                {
                    cellvalue = " ";
                }
            }
        }
        else
        {
            cellvalue = "";
        }
        return cellvalue;
    }
    
    private Row createRow(Sheet sheet, Integer rowIndex)
    {
        Row row = null;
        short minColIx = 0, maxColIx = 0;
        Map<Integer, CellStyle> styleMap = new HashMap<>();
        Row preRow = sheet.getRow(rowIndex);
        if (preRow != null)
        {
            minColIx = preRow.getFirstCellNum();
            maxColIx = preRow.getLastCellNum();
            for (short colIx = minColIx; colIx < maxColIx; colIx++)
            {
                Cell pcell = preRow.getCell(colIx);
                if (pcell == null)
                {
                    continue;
                }
                styleMap.put((int)colIx, pcell.getCellStyle());
            }
            int lastRowNo = sheet.getLastRowNum();
            sheet.shiftRows(rowIndex, lastRowNo, 1, true, true);
        }
        row = sheet.createRow(rowIndex);
        if (preRow != null)
        {
            for (short colIx = minColIx; colIx < maxColIx; colIx++)
            {
                Cell cell = row.createCell(colIx);
                cell.setCellStyle(styleMap.get((int)colIx));
            }
            
        }
        return row;
    }
}
