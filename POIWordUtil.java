import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

public class WordUtil
{
    private static final String SPLIT_FLAG = "#SPLIT_FLAG#";
    
    private static final String ROW_NUMBER = "ROW_NUMBER";
    
    public static final String BEAN_PATTERN = "\\$\\{(.*?)\\}";
    
    public static final String SUB_BEAN_PATTERN = "\\$\\{(.*?)\\}\\*\\$\\{(.*?)\\}";
    
    private static Logger LOG = LoggerFactory.getLogger(DAOUtil.class);
    
    static class CellRun
    {
        public boolean bold;
        
        public boolean italic;
        
        public String color;
        
        public String style;
        
        public String content;
        
        public String fontfmaily;
        
        public int fontsize;
        
        public int fontAlign;
        
        public CTNumPr numpr;
        
        public TextAlignment valign;
        
        public ParagraphAlignment align;
        
        public UnderlinePatterns underLine;
    }
    
    public static void parse(long dataId, JSONObject layoutJson, String path, String outpath)
    {
        if (layoutJson == null || layoutJson.isEmpty())
        {
            return;
        }
        XWPFDocument document = null;
        String mainBeanName = layoutJson.getString("bean");
        String mainBeanTitle = layoutJson.getString("title");
        long companyId = layoutJson.getLongValue("companyId");
        Map<String, String> mlableNameMap = new LinkedHashMap<>();
        Map<String, String> dateFieldMap = new LinkedHashMap<>();
        Map<String, String> referenceFieldMap = new LinkedHashMap<>();
        Map<String, Map<String, String>> stitleLableNameMap = new LinkedHashMap<>();
        try
        {
            getFieldsMap4Layout(layoutJson, mlableNameMap, referenceFieldMap, stitleLableNameMap, dateFieldMap);
            
            // 解析DOCX模板并获取document对象
            InputStream is = OSSUtil.getInstance().getFile(Constant.FLIE_LIBRARY_NAME, path);
            document = new XWPFDocument(is);
            // 处理段落
            dealParagraph(dataId, companyId, mainBeanName, mainBeanTitle, document, mlableNameMap, referenceFieldMap, stitleLableNameMap, dateFieldMap);
            // 处理表格
            dealTable(dataId, companyId, mainBeanName, mainBeanTitle, document, mlableNameMap, referenceFieldMap, stitleLableNameMap, dateFieldMap);
            // 生成新的word
            File file = new File(outpath);
            File dir = new File(file.getParent());
            if (!dir.exists())
            {
                dir.mkdirs();
            }
            file.deleteOnExit();
            FileOutputStream stream = new FileOutputStream(file);
            document.createNumbering();
            document.write(stream);
            stream.close();
        }
        catch (Exception e)
        {
            LOG.error(e.getMessage(), e);
        }
    }
    
    private static void copyParagraph(XWPFParagraph sourcePGH, XWPFParagraph targetPGH)
    {
        targetPGH.getCTP().setPPr(sourcePGH.getCTP().getPPr());
        targetPGH.setStyle(sourcePGH.getStyle());
        targetPGH.setAlignment(sourcePGH.getAlignment());
        targetPGH.setFontAlignment(sourcePGH.getFontAlignment());
        targetPGH.setVerticalAlignment(sourcePGH.getVerticalAlignment());
        if (sourcePGH.getRuns() != null)
        {
            for (XWPFRun cellR : sourcePGH.getRuns())
            {
                if (StringUtils.isNotBlank(cellR.getText(-1)))
                {
                    XWPFRun tcr = targetPGH.createRun();
                    tcr.setText(cellR.getText(-1));
                    tcr.setFontFamily(cellR.getFontFamily());
                    tcr.setFontSize(cellR.getFontSize());
                    tcr.setBold(cellR.isBold());
                    tcr.setItalic(cellR.isItalic());
                    tcr.setImprinted(cellR.isImprinted());
                    tcr.setEmbossed(cellR.isEmbossed());
                    tcr.setUnderline(cellR.getUnderline());
                }
            }
        }
    }
    
    private static void copyCell(XWPFTableCell sourceCell, XWPFTableCell targetCell)
    {
        targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
        if (sourceCell.getParagraphs() != null && sourceCell.getParagraphs().size() > 0)
        {
            for (int p = 0; p < sourceCell.getParagraphs().size(); p++)
            {
                XWPFParagraph tp = targetCell.getParagraphs().get(0);
                if (p > 0)
                {
                    tp = targetCell.addParagraph();
                }
                XWPFParagraph sp = sourceCell.getParagraphs().get(p);
                copyParagraph(sp, tp);
                if (sp.getRuns() == null)
                {
                    targetCell.setText(sourceCell.getText());
                }
            }
        }
        else
        {
            targetCell.setText(sourceCell.getText());
        }
    }
    
    private static void copyRow(XWPFTableRow sourceRow, XWPFTableRow targetRow)
    {
        if (targetRow == null)
        {
            return;
        }
        if (sourceRow.getCtRow().getTrPr() != null)
        {
            targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        }
        List<XWPFTableCell> cellList = sourceRow.getTableCells();
        if (null == cellList)
        {
            return;
        }
        
        for (XWPFTableCell sourceCell : cellList)
        {
            XWPFTableCell targetCell = targetRow.addNewTableCell();
            copyCell(sourceCell, targetCell);
        }
    }
    
    private static void dealTable(long dataId, long companyId, String mainBeanName, String mainBeanTitle, XWPFDocument document, Map<String, String> mlableNameMap,
        Map<String, String> referenceFieldMap, Map<String, Map<String, String>> stitleLableNameMap, Map<String, String> dateFieldMap)
    {
        for (XWPFTable xwpfTable : document.getTables())
        {
            boolean right = false;
            String lastSubName = null;
            List<XWPFTableRow> rows = xwpfTable.getRows();
            List<JSONObject> rowsubformJsons = new ArrayList<>();
            int rowsize=rows.size();
            Map<String, List<String>> rowBeanFieldsMap = new LinkedHashMap<>();
            for (int ri = 0; ri < rowsize; ri++)
            {
                String currentSubName = null;
                JSONObject rowSubformDataJson = new JSONObject();
                XWPFTableRow xwpfTableRow = rows.get(ri);
                out: for (int ci = 0; ci < xwpfTableRow.getTableCells().size(); ci++)
                {
                    XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(ci);
                    String cellText = xwpfTableCell.getText();
                    Map<String, List<String>> beanFieldsMap = extractor(mainBeanTitle, cellText);
                    if (!beanFieldsMap.isEmpty())
                    {
                        JSONObject subfromJson = new JSONObject();
                        for (Map.Entry<String, List<String>> entry : beanFieldsMap.entrySet())
                        {
                            String beanName = entry.getKey();
                            List<String> tfields = entry.getValue();
                            subfromJson.put(beanName, tfields);
                            String subformName = rowSubformDataJson.getString("subform");
                            if (StringUtils.isBlank(beanName) || (!beanName.equals(mainBeanTitle) && subformName != null && !subformName.equals(beanName)))
                            {
                                continue out;
                            }
                            if (!beanName.equals(mainBeanTitle))
                            {
                                currentSubName = beanName;
                            }
                            List<String> fields = rowBeanFieldsMap.get(beanName);
                            if (fields == null)
                            {
                                fields = new ArrayList<>();
                            }
                            fields.addAll(tfields);
                            rowBeanFieldsMap.put(beanName, fields);
                        }
                        JSONArray cellArr = rowSubformDataJson.getJSONArray("cells");
                        if (cellArr == null)
                        {
                            cellArr = new JSONArray();
                            rowSubformDataJson.put("cells", cellArr);
                        }
                        if (rowsubformJsons.isEmpty())
                        {
                            right = cellText.contains("*r");
                        }
                        if (currentSubName == null)
                        {
                            currentSubName = mainBeanTitle;
                        }
                        JSONObject cellJson = new JSONObject();
                        cellJson.put("index", ci);
                        cellJson.put("fields", subfromJson);
                        cellArr.add(cellJson);
                        rowSubformDataJson.put("rowIndex", ri);
                        rowSubformDataJson.put("subform", currentSubName);
                    }
                }
                
                if (currentSubName == null && lastSubName == null)
                {
                    continue;
                }
                boolean flag = true;
                if (lastSubName == null || lastSubName.equals(currentSubName))
                {
                    flag = false;
                    rowsubformJsons.add(rowSubformDataJson);
                    lastSubName = currentSubName;
                }
                if (flag || ri == rowsize- 1)
                {
                    lastSubName = currentSubName;
                    List<JSONObject> dataJsonLS =
                        dealTableSql(dataId, companyId, mainBeanName, mainBeanTitle, rowBeanFieldsMap, mlableNameMap, referenceFieldMap, stitleLableNameMap);
                    if (dataJsonLS.isEmpty())
                    {
                        continue;
                    }
                    int addRows = 0;
                    int rowSindex = rowsubformJsons.get(0).getIntValue("rowIndex");
                    int rowEindex = rowsubformJsons.get(rowsubformJsons.size() - 1).getIntValue("rowIndex");
                    int subrowSize = rowEindex - rowSindex + 1;
                    for (int i = 1; i < dataJsonLS.size(); i++)
                    {
                        JSONObject subDataJson = dataJsonLS.get(i);
                        if (!right)
                        {
                            addRows = addRows + subrowSize;
                            for (JSONObject mete : rowsubformJsons)
                            {
                                int rowIndex = mete.getIntValue("rowIndex");
                                XWPFTableRow row = rows.get(rowIndex);
                                XWPFTableRow targetRow = xwpfTable.insertNewTableRow(rowIndex+addRows);
                                copyRow(row, targetRow);
                                JSONArray cellArr = mete.getJSONArray("cells");
                                for (int ci = 0; ci < cellArr.size(); ci++)
                                {
                                    JSONObject cellJson = cellArr.getJSONObject(ci);
                                    int cindex = cellJson.getIntValue("index");
                                    JSONObject subfromJson = cellJson.getJSONObject("fields");
                                    XWPFTableCell xwpfTableCell = targetRow.getCell(cindex);
                                    for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs())
                                    {
                                        parseParagraph(cindex
                                            + 1, document, xwpfParagraph, mainBeanTitle, subfromJson, mlableNameMap, stitleLableNameMap, dateFieldMap, subDataJson);
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (JSONObject mete : rowsubformJsons)
                            {
                                int rowIndex = mete.getIntValue("rowIndex");
                                JSONArray cellArr = mete.getJSONArray("cells");
                                XWPFTableRow row = rows.get(rowIndex);
                                int cellNums = row.getTableCells().size();
                                int csi = cellArr.getJSONObject(0).getIntValue("index");
                                int csize = cellNums - csi + 1;
                                for (int t = 1; t <= (cellNums - csi); t++)
                                {
                                    XWPFTableCell targetCell = row.addNewTableCell();
                                    XWPFTableCell sourceCell = row.getCell(csi + t - 1);
                                    copyCell(sourceCell, targetCell);
                                }
                                for (int ci = 0; ci < cellArr.size(); ci++)
                                {
                                    JSONObject cellJson = cellArr.getJSONObject(ci);
                                    int cindex = cellJson.getIntValue("index");
                                    JSONObject subfromJson = cellJson.getJSONObject("fields");
                                    int nindex = (i - 1) * csize + (cindex - csi);
                                    XWPFTableCell xwpfTableCell = row.getCell(nindex);
                                    for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs())
                                    {
                                        parseParagraph(cindex
                                            + 1, document, xwpfParagraph, mainBeanTitle, subfromJson, mlableNameMap, stitleLableNameMap, dateFieldMap, subDataJson);
                                    }
                                }
                            }
                        }
                    }
                    for (JSONObject mete : rowsubformJsons)
                    {
                        int rowIndex = mete.getIntValue("rowIndex");
                        XWPFTableRow row = rows.get(rowIndex);
                        JSONArray cellArr = mete.getJSONArray("cells");
                        for (int ci = 0; ci < cellArr.size(); ci++)
                        {
                            JSONObject cellJson = cellArr.getJSONObject(ci);
                            int cindex = cellJson.getIntValue("index");
                            JSONObject subfromJson = cellJson.getJSONObject("fields");
                            XWPFTableCell xwpfTableCell = row.getCell(cindex);
                            for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs())
                            {
                                parseParagraph(cindex + 1, document, xwpfParagraph, mainBeanTitle, subfromJson, mlableNameMap, stitleLableNameMap, dateFieldMap, dataJsonLS.get(0));
                            }
                        }
                    }
                    rowsubformJsons.clear();
                    rowBeanFieldsMap.clear();
                }
            }
        }
    }
    
    private static List<JSONObject> dealTableSql(long dataId, long companyId, String mainBeanName, String mainBeanTitle, Map<String, List<String>> beanFieldsMap,
        Map<String, String> mlableNameMap, Map<String, String> referenceFieldMap, Map<String, Map<String, String>> stitleLableNameMap)
    {
        StringBuilder sqlSB = new StringBuilder();
        StringBuilder selectFieldSB = new StringBuilder();
        Set<String> beans = beanFieldsMap.keySet();
        List<JSONObject> dataJsonLS = new ArrayList<>();
        String employeeTable = DAOUtil.getTableName(TableConstant.TABLE_EMPLOYEE, companyId);
        String departmentTable = DAOUtil.getTableName(TableConstant.TABLE_DEPARTMENT, companyId);
        String attachmentTable = DAOUtil.getTableName(TableConstant.TABLE_ATTACHMENT, companyId);
        String fromSql = dealFromSql(dataId, companyId, mainBeanName, mainBeanTitle, beans, mlableNameMap);
        if (StringUtils.isBlank(fromSql))
        {
            return dataJsonLS;
        }
        for (Map.Entry<String, List<String>> entry : beanFieldsMap.entrySet())
        {
            String alias = "m.";
            String beanTitle = entry.getKey();
            String currentBean = mainBeanName;
            if (!beanTitle.equals(mainBeanTitle))
            {
                alias = "s.";
                currentBean = mainBeanName.concat("_").concat(mlableNameMap.get(beanTitle));
            }
            Map<String, String> lableNameMap = stitleLableNameMap.get(beanTitle);
            if (lableNameMap == null)
            {
                lableNameMap = mlableNameMap;
            }
            if (lableNameMap != null)
            {
                List<String> labels = new ArrayList<>();
                for (String lable : entry.getValue())
                {
                    if (labels.contains(lable))
                    {
                        continue;
                    }
                    labels.add(lable);
                    String field = lableNameMap.get(lable);
                    if (field == null)
                    {
                        continue;
                    }
                    String tfield = alias.concat(field);
                    if (selectFieldSB.length() > 0)
                    {
                        selectFieldSB.append(",");
                    }
                    if (tfield.contains(TableConstant.TYPE_REFERENCE))
                    {
                        String tvalue = referenceFieldMap.get(lable);
                        String[] tarr = tvalue.split(SPLIT_FLAG);
                        if (tarr.length >= 3)
                        {
                            String rfbean = tarr[1];
                            String rffield = tarr[2];
                            String rtable = DAOUtil.getTableName(rfbean, companyId);
                            selectFieldSB.append("(select ")
                                .append(rffield)
                                .append(" from ")
                                .append(rtable)
                                .append(" where id=")
                                .append(alias)
                                .append(field)
                                .append(")as ")
                                .append(field);
                        }
                    }
                    else if (field.contains(TableConstant.TYPE_PICTURE) || field.contains(TableConstant.TYPE_ATTACHMENT) || field.contains(TableConstant.TYPE_SIGNATURE))
                    {
                        selectFieldSB.append("(selectstring_agg(file_url,',')")
                            .append(" from ")
                            .append(attachmentTable)
                            .append(" where data_id=")
                            .append(dataId)
                            .append(" and bean='")
                            .append(currentBean)
                            .append("' and original_file_name='")
                            .append(field)
                            .append("' and ")
                            .append(TableConstant.FIELD_DEL_STATUS)
                            .append("=")
                            .append(TableConstant.DEL_STATUS_NORMAL)
                            .append(")as ")
                            .append(field);
                    }
                    else if (JSONParser4SQLNew.checkVL(tfield, null))
                    {
                        selectFieldSB.append(tfield).append(TableConstant.PICKUP_LABEL_FIELD_SUFFIX).append(" as ").append(field);
                    }
                    else if (tfield.contains(TableConstant.TYPE_PERSONNEL))
                    {
                        selectFieldSB.append("(select string_agg(")
                            .append(TableConstant.FIELD_EMPLOYEE_NAME)
                            .append(",',') from ")
                            .append(employeeTable)
                            .append(" where array[id]<@string_to_array(")
                            .append(tfield)
                            .append("::varchar,',')::int[] and ")
                            .append(TableConstant.FIELD_DEL_STATUS)
                            .append("=")
                            .append(TableConstant.DEL_STATUS_NORMAL)
                            .append(") as ")
                            .append(field);
                    }
                    else if (tfield.contains(TableConstant.TYPE_DEPARTMENT))
                    {
                        selectFieldSB.append("(select string_agg(")
                            .append(TableConstant.FIELD_DEPARTMENT_NAME)
                            .append(",',') from ")
                            .append(departmentTable)
                            .append(" where array[id]<@string_to_array(")
                            .append(tfield)
                            .append("::varchar,',')::int[] and ")
                            .append(TableConstant.FIELD_DEL_STATUS)
                            .append("=")
                            .append(TableConstant.DEL_STATUS_NORMAL)
                            .append(") as ")
                            .append(field);
                    }
                    else
                    {
                        selectFieldSB.append(tfield);
                    }
                }
            }
        }
        if (selectFieldSB.length() > 0)
        {
            sqlSB.append("select ").append(selectFieldSB).append(fromSql);
        }
        if (sqlSB.length() > 0)
        {
            dataJsonLS = DAOUtil.executeQuery4JSON(sqlSB.toString(), new ArrayList<>());
        }
        return dataJsonLS;
    }
    
    private static String dealFromSql(long dataId, long companyId, String mainBeanName, String mainBeanTitle, Set<String> beanTitles, Map<String, String> mlableNameMap)
    {
        StringBuilder fromSB = new StringBuilder();
        if (beanTitles.size() == 1)
        {
            String alias = " m ";
            String idfield = "id";
            String beanName = mainBeanName;
            String beanTitle = beanTitles.iterator().next();
            if (!beanTitle.equals(mainBeanTitle))
            {
                alias = " s ";
                idfield = mainBeanName.concat("_id");
                String fieldName = mlableNameMap.get(beanTitle);
                if (!fieldName.contains(TableConstant.TYPE_SUBFORM))
                {
                    return null;
                }
                beanName = mainBeanName.concat("_").concat(fieldName);
            }
            String table = DAOUtil.getTableName(beanName, companyId);
            fromSB.append(" from ")
                .append(table)
                .append(" ")
                .append(alias)
                .append(" where ")
                .append(idfield)
                .append("=")
                .append(dataId)
                .append(" and ")
                .append(TableConstant.FIELD_DEL_STATUS)
                .append("=")
                .append(TableConstant.DEL_STATUS_NORMAL)
                .append(" order by id");
        }
        else if (beanTitles.size() == 2)
        {
            StringBuilder tempSB = new StringBuilder();
            for (String beanTitle : beanTitles)
            {
                String alias = " m ";
                String beanName = mainBeanName;
                if (!beanTitle.equals(mainBeanTitle))
                {
                    alias = " s ";
                    String fieldName = mlableNameMap.get(beanTitle);
                    if (!fieldName.contains(TableConstant.TYPE_SUBFORM))
                    {
                        return null;
                    }
                    beanName = mainBeanName.concat("_").concat(fieldName);
                }
                String table = DAOUtil.getTableName(beanName, companyId);
                if (tempSB.length() > 0)
                {
                    tempSB.append(" join ");
                }
                tempSB.append(table).append(alias);
            }
            fromSB.append(" from ")
                .append(tempSB)
                .append(" on m.id=s.")
                .append(mainBeanName)
                .append("_id and s.")
                .append(TableConstant.FIELD_DEL_STATUS)
                .append("=")
                .append(TableConstant.DEL_STATUS_NORMAL)
                .append(" where m.id=")
                .append(dataId)
                .append(" order by s.id");
        }
        else
        {
            return null;
        }
        
        return fromSB.toString();
    }
    
    public static void dealParagraph(long dataId, long companyId, String mainBeanName, String mainBeanTitle, XWPFDocument document, Map<String, String> mlableNameMap,
        Map<String, String> referenceFieldMap, Map<String, Map<String, String>> stitleLableNameMap, Map<String, String> dateFieldMap)
    {
        String lastSubName = null;
        JSONObject subfromJson = new JSONObject();
        List<XWPFParagraph> paragraphs = new ArrayList<>();
        Map<String, List<String>> rowBeanFieldsMap = new LinkedHashMap<>();
        List<XWPFParagraph> xwpfParagraphs = document.getParagraphs();
        int pgsize = xwpfParagraphs.size();
        out: for (int p = 0; p <pgsize; p++)
        {
            XWPFParagraph xwpfParagraph = xwpfParagraphs.get(p);
            String currentSubName = null;
            JSONObject rowSubformDataJson = new JSONObject();
            String content = xwpfParagraph.getText();
            Map<String, List<String>> beanFieldsMap = extractor(mainBeanTitle, content);
            if (!beanFieldsMap.isEmpty())
            {
                for (Map.Entry<String, List<String>> entry : beanFieldsMap.entrySet())
                {
                    String beanName = entry.getKey();
                    List<String> tfields = entry.getValue();
                    subfromJson.put(beanName, tfields);
                    String subformName = rowSubformDataJson.getString("subform");
                    if (StringUtils.isBlank(beanName) || (!beanName.equals(mainBeanTitle) && subformName != null && !subformName.equals(beanName)))
                    {
                        continue out;
                    }
                    if (!beanName.equals(mainBeanTitle))
                    {
                        currentSubName = beanName;
                    }
                    List<String> fields = rowBeanFieldsMap.get(beanName);
                    if (fields == null)
                    {
                        fields = new ArrayList<>();
                    }
                    fields.addAll(tfields);
                    rowBeanFieldsMap.put(beanName, fields);
                }
                if (currentSubName == null)
                {
                    currentSubName = mainBeanTitle;
                }
                rowSubformDataJson.put("subform", currentSubName);
            }
            if (currentSubName == null && lastSubName == null)
            {
                continue;
            }
            boolean flag = true;
            if (lastSubName == null || lastSubName.equals(currentSubName))
            {
                flag = false;
                paragraphs.add(xwpfParagraph);
                lastSubName = currentSubName;
            }
            if (flag || p ==pgsize - 1)
            {
                lastSubName = currentSubName;
                List<JSONObject> dataJsonLS = dealTableSql(dataId, companyId, mainBeanName, mainBeanTitle, rowBeanFieldsMap, mlableNameMap, referenceFieldMap, stitleLableNameMap);
                if (dataJsonLS.isEmpty())
                {
                    continue;
                }
                for (int i = 1; i < dataJsonLS.size(); i++)
                {
                    JSONObject subDataJson = dataJsonLS.get(i);
                    for (XWPFParagraph paragraph : paragraphs)
                    {
                        XmlCursor cursor = paragraph.getCTP().newCursor();
                        XWPFParagraph newPara = document.insertNewParagraph(cursor);
                        copyParagraph(paragraph, newPara);
                        parseParagraph(0, document, newPara, mainBeanTitle, subfromJson, mlableNameMap, stitleLableNameMap, dateFieldMap, subDataJson);
                    }
                }
                for (XWPFParagraph paragraph : paragraphs)
                {
                    parseParagraph(0, document, paragraph, mainBeanTitle, subfromJson, mlableNameMap, stitleLableNameMap, dateFieldMap, dataJsonLS.get(0));
                }
                subfromJson.clear();
                paragraphs.clear();
                rowBeanFieldsMap.clear();
            }
        }
    }
    
    public static void dealParagraph(long dataId, long companyId, String mainBeanName, String mainBeanTitle, XWPFDocument document, Map<String, String> mlableNameMap,
        Map<String, String> referenceFieldMap, Map<String, String> dateFieldMap)
    {
        Set<String> beans = new HashSet<>();
        Map<String, Set<String>> allBeanFieldsMap = new LinkedHashMap<>();
        Map<XWPFParagraph, Map<String, List<String>>> paragraphKeywordMap = new LinkedHashMap<>();
        for (XWPFParagraph xwpfParagraph : document.getParagraphs())
        {
            String content = xwpfParagraph.getText();
            Map<String, List<String>> beanFieldsMap = extractor(mainBeanTitle, content);
            if (!beanFieldsMap.isEmpty())
            {
                for (Map.Entry<String, List<String>> entry : beanFieldsMap.entrySet())
                {
                    String beanName = entry.getKey();
                    Set<String> fields = allBeanFieldsMap.get(beanName);
                    if (fields == null)
                    {
                        fields = new HashSet<>();
                    }
                    beans.add(beanName);
                    fields.addAll(entry.getValue());
                    allBeanFieldsMap.put(beanName, fields);
                }
                paragraphKeywordMap.put(xwpfParagraph, beanFieldsMap);
            }
        }
        // 组装查询获取对应数据
        String employeeTable = DAOUtil.getTableName(TableConstant.TABLE_EMPLOYEE, companyId);
        String departmentTable = DAOUtil.getTableName(TableConstant.TABLE_DEPARTMENT, companyId);
        String attachmentTable = DAOUtil.getTableName(TableConstant.TABLE_ATTACHMENT, companyId);
        StringBuilder sqlSB = new StringBuilder();
        if (beans.contains(mainBeanTitle) && beans.size() == 1)
        {
            String table = DAOUtil.getTableName(mainBeanName, companyId);
            Set<String> lables = allBeanFieldsMap.get(mainBeanTitle);
            sqlSB.append("select ");
            for (String lable : lables)
            {
                String field = mlableNameMap.get(lable);
                if (StringUtils.isBlank(field))
                {
                    continue;
                }
                if (field.contains(TableConstant.TYPE_REFERENCE))
                {
                    String tvalue = referenceFieldMap.get(lable);
                    String[] tarr = tvalue.split(SPLIT_FLAG);
                    if (tarr.length >= 3)
                    {
                        String rfbean = tarr[1];
                        String rffield = tarr[2];
                        String rtable = DAOUtil.getTableName(rfbean, companyId);
                        sqlSB.append("(select ").append(rffield).append(" from ").append(rtable).append(" where id=m.").append(field).append(")as ").append(field).append(",");
                    }
                }
                else if (field.contains(TableConstant.TYPE_PICTURE) || field.contains(TableConstant.TYPE_ATTACHMENT) || field.contains(TableConstant.TYPE_SIGNATURE))
                {
                    sqlSB.append("(select string_agg(file_url,',')")
                        .append(" from ")
                        .append(attachmentTable)
                        .append(" where data_id=")
                        .append(dataId)
                        .append(" and bean='")
                        .append(mainBeanName)
                        .append("' and original_file_name='")
                        .append(field)
                        .append("' and ")
                        .append(TableConstant.FIELD_DEL_STATUS)
                        .append("=")
                        .append(TableConstant.DEL_STATUS_NORMAL)
                        .append(")as ")
                        .append(field)
                        .append(",");
                }
                else if (JSONParser4SQLNew.checkVL(field, null))
                {
                    sqlSB.append(field).append(TableConstant.PICKUP_LABEL_FIELD_SUFFIX).append(",");
                }
                else if (field.contains(TableConstant.TYPE_PERSONNEL))
                {
                    sqlSB.append("(select string_agg(")
                        .append(TableConstant.FIELD_EMPLOYEE_NAME)
                        .append(",',') from ")
                        .append(employeeTable)
                        .append(" where array[id]<@string_to_array(m.")
                        .append(field)
                        .append("::varchar,',')::int[] and ")
                        .append(TableConstant.FIELD_DEL_STATUS)
                        .append("=")
                        .append(TableConstant.DEL_STATUS_NORMAL)
                        .append(") as ")
                        .append(field)
                        .append(",");
                }
                else if (field.contains(TableConstant.TYPE_DEPARTMENT))
                {
                    sqlSB.append("(select string_agg(")
                        .append(TableConstant.FIELD_DEPARTMENT_NAME)
                        .append(",',') from ")
                        .append(departmentTable)
                        .append(" where array[id]<@string_to_array(m.")
                        .append(field)
                        .append("::varchar,',')::int[] and ")
                        .append(TableConstant.FIELD_DEL_STATUS)
                        .append("=")
                        .append(TableConstant.DEL_STATUS_NORMAL)
                        .append(") as ")
                        .append(field)
                        .append(",");
                }
                else
                {
                    sqlSB.append(field).append(",");
                }
            }
            sqlSB.append("'logo' as logo from ").append(table).append(" m where id=").append(dataId);
            JSONObject dataJson = DAOUtil.executeQuery4FirstJSON(sqlSB.toString(), new ArrayList<>());
            for (Map.Entry<XWPFParagraph, Map<String, List<String>>> entry : paragraphKeywordMap.entrySet())
            {
                parseParagraph(0, document, entry.getKey(), mainBeanTitle, entry.getValue(), mlableNameMap, null, dateFieldMap, dataJson, true, true);
            }
        }
    }
    
    private static CellRun parseParagraph(int rowNum, XWPFDocument document, XWPFParagraph paragraph, String mainBeanTitle, Map<String, List<String>> matchesMap,
        Map<String, String> mlableNameMap, Map<String, Map<String, String>> stitleLableNameMap, Map<String, String> dateFieldMap, JSONObject dataJson, boolean fixLength,
        boolean removePargraph)
    {
        
        int fontsize = 0;
        String color = null;
        String fontfmaily = null;
        boolean bold = false;
        boolean italic = false;
        CellRun cellrun = new CellRun();
        UnderlinePatterns underLine = null;
        String content = paragraph.getText();
        cellrun.content = content;
        cellrun.style = paragraph.getStyle();
        cellrun.align = paragraph.getAlignment();
        cellrun.fontAlign = paragraph.getFontAlignment();
        cellrun.valign = paragraph.getVerticalAlignment();
        List<XWPFRun> runs = paragraph.getRuns();
        int size = runs.size();
        if (size > 0)
        {
            XWPFRun tmprun = runs.get(0);
            italic = tmprun.isItalic();
            bold = tmprun.isBold();
            color = tmprun.getColor();
            underLine = tmprun.getUnderline();
            fontfmaily = tmprun.getFontFamily();
            fontsize = tmprun.getFontSize();
        }
        for (int i = 0; i < size; i++)
        {
            paragraph.removeRun(0);
        }
        try
        {
            cellrun.numpr = paragraph.getCTP().getPPr().getNumPr();
        }
        catch (Exception e)
        {
            
        }
        XWPFRun run = paragraph.createRun();
        StringBuilder tempSB = new StringBuilder();
        Map<String, String> matchFieldMap = new LinkedHashMap<>();
        
        for (Map.Entry<String, List<String>> entry : matchesMap.entrySet())
        {
            String beanTitle = entry.getKey();
            for (String match : entry.getValue())
            {
                String rowNumField = "${".concat(ROW_NUMBER).concat("}");
                if (match.equals(ROW_NUMBER))
                {
                    content = content.replace(rowNumField, String.valueOf(rowNum));
                    continue;
                }
                String field = mlableNameMap.get(match);
                String oldmatch = "${".concat(match).concat("}");
                if (!beanTitle.equals(mainBeanTitle))
                {
                    if (stitleLableNameMap == null || stitleLableNameMap.get(beanTitle) == null)
                    {
                        continue;
                    }
                    oldmatch = "${".concat(beanTitle).concat("}*${").concat(match).concat("}");
                    Map<String, String> subLableNameMap = stitleLableNameMap.get(beanTitle);
                    field = subLableNameMap.get(match);
                }
                
                String tempv = dataJson.getString(field);
                if (StringUtils.isBlank(tempv))
                {
                    content = content.replace(oldmatch, "");
                    continue;
                }
                if (field.contains(TableConstant.TYPE_PICTURE) || field.contains(TableConstant.TYPE_ATTACHMENT) || field.contains(TableConstant.TYPE_SIGNATURE))
                {
                    matchFieldMap.put(oldmatch, field);
                }
                else
                {
                    try
                    {
                        if (dateFieldMap.containsKey(field))
                        {
                            SimpleDateFormat sdf = new SimpleDateFormat(dateFieldMap.get(field));
                            tempv = sdf.format(new Date(Long.parseLong(tempv)));
                        }
                        else if (JSONParser4SQLNew.checkVL(field, null))
                        {
                            tempv = tempv.replace("{", "").replace("}", "");
                        }
                        else if (field.contains(TableConstant.TYPE_LOCATION))
                        {
                            JSONObject tjson = JSONObject.parseObject(tempv);
                            tempv = tjson.getString("value");
                        }
                        else if (field.contains(TableConstant.TYPE_MULTITEXT))
                        {
                            tempv = MultiTextUtil.getContent(tempv);
                        }
                        else if (!field.contains(TableConstant.TYPE_TEXTAREA) && field.contains(TableConstant.TYPE_AREA))
                        {
                            tempSB.setLength(0);
                            for (String tv : tempv.split(","))
                            {
                                if (tempSB.length() > 0)
                                {
                                    tempSB.append(" ");
                                }
                                tempSB.append(tv.split(":")[1]);
                            }
                            tempv = tempSB.toString();
                        }
                    }
                    catch (Exception e)
                    {
                        LOG.error(e.getMessage(), e);
                    }
                    
                    String tempMatch = oldmatch;
                    int nowlength = tempv.length();
                    int oldlength = tempMatch.length();
                    StringBuilder spaceSB = new StringBuilder(tempv);
                    if (fixLength)
                    {
                        for (int i = 1; i <= (oldlength - nowlength); i++)
                        {
                            spaceSB.append(" ");
                        }
                    }
                    content = content.replace(tempMatch, spaceSB.toString());
                }
            }
        }
        run.setBold(bold);
        cellrun.bold = bold;
        run.setItalic(italic);
        cellrun.italic = italic;
        if (!matchFieldMap.isEmpty())
        {
            int lastIndex = 0;
            for (Map.Entry<String, String> entry : matchFieldMap.entrySet())
            {
                String match = entry.getKey();
                String field = entry.getValue();
                String tempv = dataJson.getString(field);
                int index = content.indexOf(match);
                if (index > 0 && lastIndex < content.length())
                {
                    run.setText(content.substring(lastIndex, index));
                }
                lastIndex = index + match.length();
                for (String pictureUrl : tempv.split(","))
                {
                    dealPicture(document, run, pictureUrl, field);
                }
            }
        }
        else if (content.contains("\n"))
        {
            String[] lines = content.split("\n");
            for (int i = 0; i < lines.length; i++)
            {
                if (i > 0)
                {
                    run.addBreak();
                }
                run.setText(lines[i]);
            }
        }
        else
        {
            run.setText(content);
        }
        if (removePargraph && StringUtils.isBlank(content))
        {
            document.removeBodyElement(document.getPosOfParagraph(paragraph));
            return null;
        }
        if (StringUtils.isNotBlank(color))
        {
            run.setColor(color);
            cellrun.color = color;
        }
        if (underLine != null)
        {
            run.setUnderline(underLine);
            cellrun.underLine = underLine;
        }
        if (fontsize > 0)
        {
            cellrun.fontsize = fontsize;
            run.setFontSize(fontsize);
        }
        if (fontfmaily != null)
        {
            cellrun.fontfmaily = fontfmaily;
            run.setFontFamily(fontfmaily);
        }
        paragraph.addRun(run);
        return cellrun;
    }
    
    public static void parseParagraph(int rowNum, XWPFDocument document, XWPFParagraph paragraph, String mainBeanTitle, JSONObject subformJson, Map<String, String> mlableNameMap,
        Map<String, Map<String, String>> stitleLableNameMap, Map<String, String> dateFieldMap, JSONObject dataJson)
    {
        String content = paragraph.getText();
        List<XWPFRun> runs = paragraph.getRuns();
        int size = runs.size();
        XWPFRun frun = paragraph.createRun();
        if (size > 0)
        {
            XWPFRun cellR = runs.get(0);
            frun = paragraph.createRun();
            frun.setFontFamily(cellR.getFontFamily());
            frun.setFontSize(cellR.getFontSize());
            frun.setBold(cellR.isBold());
            frun.setItalic(cellR.isItalic());
            frun.setImprinted(cellR.isImprinted());
            frun.setEmbossed(cellR.isEmbossed());
            frun.setUnderline(cellR.getUnderline());
            for (int i = 0; i < size; i++)
            {
                paragraph.removeRun(0);
            }
        }
        StringBuilder tempSB = new StringBuilder();
        Map<String, String> matchFieldMap = new LinkedHashMap<>();
        for (String beanTitle : subformJson.keySet())
        {
            JSONArray arr = subformJson.getJSONArray(beanTitle);
            for (Object obj : arr)
            {
                String match = obj.toString();
                
                String rowNumField = "${".concat(ROW_NUMBER).concat("}");
                if (match.equals(ROW_NUMBER))
                {
                    content = content.replace(rowNumField, String.valueOf(rowNum));
                    continue;
                }
                String field = mlableNameMap.get(match);
                String oldmatch = "${".concat(match).concat("}");
                if (!beanTitle.equals(mainBeanTitle))
                {
                    if (stitleLableNameMap == null || stitleLableNameMap.get(beanTitle) == null)
                    {
                        continue;
                    }
                    oldmatch = "${".concat(beanTitle).concat("}*${").concat(match).concat("}");
                    Map<String, String> subLableNameMap = stitleLableNameMap.get(beanTitle);
                    field = subLableNameMap.get(match);
                }
                
                String tempv = dataJson.getString(field);
                if (StringUtils.isBlank(tempv))
                {
                    content = content.replace(oldmatch, "");
                    continue;
                }
                if (field.contains(TableConstant.TYPE_PICTURE) || field.contains(TableConstant.TYPE_ATTACHMENT) || field.contains(TableConstant.TYPE_SIGNATURE))
                {
                    matchFieldMap.put(oldmatch, field);
                }
                else
                {
                    try
                    {
                        if (dateFieldMap.containsKey(field))
                        {
                            SimpleDateFormat sdf = new SimpleDateFormat(dateFieldMap.get(field));
                            tempv = sdf.format(new Date(Long.parseLong(tempv)));
                        }
                        else if (JSONParser4SQLNew.checkVL(field, null))
                        {
                            tempv = tempv.replace("{", "").replace("}", "");
                        }
                        else if (field.contains(TableConstant.TYPE_LOCATION))
                        {
                            JSONObject tjson = JSONObject.parseObject(tempv);
                            tempv = tjson.getString("value");
                        }
                        else if (field.contains(TableConstant.TYPE_MULTITEXT))
                        {
                            tempv = MultiTextUtil.getContent(tempv);
                        }
                        else if (!field.contains(TableConstant.TYPE_TEXTAREA) && field.contains(TableConstant.TYPE_AREA))
                        {
                            tempSB.setLength(0);
                            for (String tv : tempv.split(","))
                            {
                                if (tempSB.length() > 0)
                                {
                                    tempSB.append(" ");
                                }
                                tempSB.append(tv.split(":")[1]);
                            }
                            tempv = tempSB.toString();
                        }
                    }
                    catch (Exception e)
                    {
                        LOG.error(e.getMessage(), e);
                    }
                    String tempMatch = oldmatch;
                    StringBuilder spaceSB = new StringBuilder(tempv);
                    content = content.replace(tempMatch, spaceSB.toString());
                }
            }
        }
        if (!matchFieldMap.isEmpty())
        {
            int lastIndex = 0;
            for (Map.Entry<String, String> entry : matchFieldMap.entrySet())
            {
                String match = entry.getKey();
                String field = entry.getValue();
                String tempv = dataJson.getString(field);
                int index = content.indexOf(match);
                if (index > 0 && lastIndex < content.length())
                {
                    frun.setText(content.substring(lastIndex, index));
                }
                lastIndex = index + match.length();
                for (String pictureUrl : tempv.split(","))
                {
                    dealPicture(document, frun, pictureUrl, field);
                }
            }
        }
        else if (content.contains("\n"))
        {
            String[] lines = content.split("\n");
            for (int i = 0; i < lines.length; i++)
            {
                if (i > 0)
                {
                    frun.addBreak();
                }
                frun.setText(lines[i]);
            }
        }
        else
        {
            frun.setText(content);
        }
    }
    
    public static void dealPicture(XWPFDocument document, XWPFRun run, String value, String field)
    {
        try
        {
            InputStream is = null;
            String imgtype = value.substring(value.lastIndexOf(".") + 1).toLowerCase();
            int format = XWPFDocument.PICTURE_TYPE_JPEG;
            if (imgtype.equals("png"))
            {
                format = XWPFDocument.PICTURE_TYPE_PNG;
            }
            if (value.startsWith("/common/file/download"))
            {
                String buckName = Constant.FLIE_LIBRARY_NAME;
                String fileName = value.substring(value.indexOf("fileName=") + "fileName=".length());
                is = OSSUtil.getInstance().getFile(buckName, fileName);
            }
            else
            {
                is = new FileInputStream(value);
            }
            byte[] oldByts = new byte[is.available()];
            is.read(oldByts);
            is.close();
            byte[] oldBytsBack = oldByts.clone();
            BufferedImage src = javax.imageio.ImageIO.read(new ByteArrayInputStream(oldByts));
            String picId = document.addPictureData(new ByteArrayInputStream(oldBytsBack), format);
            int width = src.getWidth() > 600 ? 600 : src.getWidth();
            int height = src.getHeight() > 800 ? 800 : src.getHeight();
            if (field.contains(TableConstant.TYPE_SIGNATURE))
            {
                width = 108;
                height = 60;
            }
            createPicture(run, picId, document.getNextPicNameNumber(format), width, height);
        }
        catch (Exception e)
        {
            LOG.error(e.getMessage(), e);
        }
    }
    
    public static Map<String, List<String>> extractor(String mainBeanTitle, String content)
    {
        Map<String, List<String>> beanFieldsMap = new LinkedHashMap<>();
        if (StringUtils.isBlank(content))
        {
            return beanFieldsMap;
        }
        Pattern pattern = Pattern.compile(BEAN_PATTERN);
        Pattern pattern2 = Pattern.compile(SUB_BEAN_PATTERN);
        Matcher m = pattern.matcher(content);
        Matcher m2 = pattern2.matcher(content);
        List<String> values = new ArrayList<>();
        String fieldName = null;
        String beanName = null;
        while (m2.find())
        {
            beanName = m2.group(1);
            fieldName = m2.group(2);
            if (fieldName != null && beanName != null)
            {
                List<String> fields = beanFieldsMap.get(beanName);
                if (fields == null)
                {
                    fields = new ArrayList<>();
                }
                fields.add(fieldName);
                beanFieldsMap.put(beanName, fields);
            }
        }
        if (fieldName == null && beanName == null)
        {
            while (m.find())
            {
                fieldName = m.group(1);
                values.add(fieldName);
                List<String> fields = beanFieldsMap.get(mainBeanTitle);
                if (fields == null)
                {
                    fields = new ArrayList<>();
                }
                fields.add(fieldName);
                beanFieldsMap.put(mainBeanTitle, fields);
            }
        }
        
        return beanFieldsMap;
    }
    
    public static void createPicture(XWPFRun run, String blipId, int id, int width, int height)
    {
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        CTInline inline = run.getCTR().addNewDrawing().addNewInline();
        StringBuilder picxmlSB = new StringBuilder();
        picxmlSB.append("<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">")
            .append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">")
            .append("<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">")
            .append("<pic:nvPicPr><pic:cNvPr id=\"")
            .append(id)
            .append("\" name=\"Generated\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"")
            .append(blipId)
            .append("\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/><a:stretch><a:fillRect/>")
            .append("</a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>")
            .append("<a:ext cx=\"")
            .append(width)
            .append("\" cy=\"")
            .append(height)
            .append("\"/></a:xfrm><a:prstGeom prst=\"rect\">")
            .append("<a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>");
        XmlToken xmlToken = null;
        try
        {
            xmlToken = XmlToken.Factory.parse(picxmlSB.toString());
        }
        catch (XmlException xe)
        {
            LOG.error(xe.getMessage(), xe);
        }
        inline.set(xmlToken);
        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);
        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);
        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr("Generated");
    }
    
    public static void getFieldsMap4Layout(JSONObject layoutJson, Map<String, String> mlableNameMap, Map<String, String> referenceFieldMap,
        Map<String, Map<String, String>> stitleLableNameMap, Map<String, String> dateFieldMap)
    {
        StringBuilder tempSB = new StringBuilder();
        if (layoutJson != null && !layoutJson.isEmpty())
        {
            JSONArray layoutArray = layoutJson.getJSONArray("layout");
            if (layoutArray != null)
            {
                for (Object object : layoutArray)
                {
                    JSONObject json = (JSONObject)object;
                    JSONArray rows = json.getJSONArray("rows");
                    if (rows != null)
                    {
                        for (Object fieldObject : rows)
                        {
                            JSONObject fieldJson = (JSONObject)fieldObject;
                            String fieldType = fieldJson.getString("type");
                            String fieldName = fieldJson.getString("name");
                            String fieldLabel = fieldJson.getString("label");
                            JSONObject tfjson = fieldJson.getJSONObject("field");
                            mlableNameMap.put(fieldLabel, fieldName);
                            if (fieldType.equals(TableConstant.TYPE_DATETIME))
                            {
                                if (tfjson != null)
                                {
                                    dateFieldMap.put(fieldName, tfjson.getString("formatType"));
                                }
                            }
                            else if (fieldType.equals(TableConstant.TYPE_FUNCTIONFORMULA))
                            {
                                if (tfjson != null)
                                {
                                    int numberType = tfjson.getIntValue("numberType");
                                    if (numberType == AllEnum.FormulaReturnEnum.DATE.ordinal())
                                    {
                                        dateFieldMap.put(fieldName, tfjson.getString("chooseType"));
                                    }
                                }
                            }
                            if (fieldType.equals(TableConstant.TYPE_REFERENCE))
                            {
                                JSONObject rfieldJson = fieldJson.getJSONObject("relevanceField");
                                JSONObject rmoduleJson = fieldJson.getJSONObject("relevanceModule");
                                if (rmoduleJson != null && rfieldJson != null)
                                {
                                    tempSB.setLength(0);
                                    tempSB.append(rmoduleJson.getString("moduleLabel"))
                                        .append(SPLIT_FLAG)
                                        .append(rmoduleJson.getString("moduleName"))
                                        .append(SPLIT_FLAG)
                                        .append(rfieldJson.getString("fieldName"));
                                    
                                    referenceFieldMap.put(fieldLabel, tempSB.toString());
                                }
                            }
                            else if (fieldType.equals(TableConstant.TYPE_SUBFORM))
                            {
                                Map<String, String> slableNameMap = new HashMap<>();
                                JSONArray componentList = fieldJson.getJSONArray("componentList");
                                for (Object subObj : componentList)
                                {
                                    JSONObject subJson = (JSONObject)subObj;
                                    String subFieldName = subJson.getString("name");
                                    String subfieldLabel = subJson.getString("label");
                                    String subfieldType = subJson.getString("type");
                                    JSONObject subtfjson = subJson.getJSONObject("field");
                                    if (subfieldType.equals(TableConstant.TYPE_REFERENCE))
                                    {
                                        JSONObject rfieldJson = subJson.getJSONObject("relevanceField");
                                        JSONObject rmoduleJson = subJson.getJSONObject("relevanceModule");
                                        if (rmoduleJson != null && rfieldJson != null)
                                        {
                                            tempSB.setLength(0);
                                            tempSB.append(rmoduleJson.getString("moduleLabel"))
                                                .append(SPLIT_FLAG)
                                                .append(rmoduleJson.getString("moduleName"))
                                                .append(SPLIT_FLAG)
                                                .append(rfieldJson.getString("fieldName"));
                                            
                                            referenceFieldMap.put(subfieldLabel, tempSB.toString());
                                        }
                                    }
                                    if (subfieldType.equals(TableConstant.TYPE_DATETIME))
                                    {
                                        if (subtfjson != null)
                                        {
                                            dateFieldMap.put(subFieldName, subtfjson.getString("formatType"));
                                        }
                                    }
                                    else if (subfieldType.equals(TableConstant.TYPE_FUNCTIONFORMULA))
                                    {
                                        if (subtfjson != null)
                                        {
                                            int numberType = subtfjson.getIntValue("numberType");
                                            if (numberType == AllEnum.FormulaReturnEnum.DATE.ordinal())
                                            {
                                                dateFieldMap.put(subFieldName, subtfjson.getString("chooseType"));
                                            }
                                        }
                                    }
                                    slableNameMap.put(subfieldLabel, subFieldName);
                                }
                                stitleLableNameMap.put(fieldLabel, slableNameMap);
                            }
                        }
                    }
                }
            }
        }
    }
}
