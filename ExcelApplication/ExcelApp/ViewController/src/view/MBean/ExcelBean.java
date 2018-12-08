package view.MBean;


import java.text.SimpleDateFormat;


import java.util.ArrayList;
import java.util.HashMap;


import java.util.Iterator;
import java.util.TreeSet;

import java.util.regex.Pattern;

import javax.faces.component.UIComponent;
import javax.faces.event.ValueChangeEvent;

import model.java.EODExcel;

import oracle.adf.model.BindingContext;
import oracle.adf.model.binding.DCBindingContainer;


import oracle.adf.view.rich.component.rich.data.RichTable;
import oracle.adf.view.rich.component.rich.input.RichInputFile;
import oracle.adf.view.rich.component.rich.layout.RichPanelFormLayout;
import oracle.adf.view.rich.component.rich.layout.RichPanelStretchLayout;
import oracle.adf.view.rich.component.rich.output.RichMessage;
import oracle.adf.view.rich.context.AdfFacesContext;

import oracle.binding.BindingContainer;

import oracle.jbo.AttributeDef;
import oracle.jbo.ViewObject;


import org.apache.myfaces.trinidad.model.UploadedFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelBean {
    private RichTable excelTable;
    private HashMap<String, String> formatMap = new HashMap<String, String>();
    private HashMap<String, String> extensionMap = new HashMap<String, String>();
    private final String EXCEL_VIEW_ITERATOR = "ExcelViewIterator";
    private final String XLSX_FORMAT_1 = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private final String XLSX_FORMAT_2 = "application/xlsx";
    private final String XLS_FORMAT = "application/vnd.ms-excel";
    private final String XLSX = "XLSX";
    private final String XLS = "XLS";
    private final String HEX_STRING_BLUE_XLS = "0:CCCC:FFFF";
    private final String HEX_STRING_GREEN_XLS = "9999:CCCC:0";
    private final String HEX_STRING_YELLOW_XLS = "FFFF:FFFF:0";
    private final String HEX_STRING_BLUE_XLSX = "#00B0F0";
    private final String HEX_STRING_GREEN_XLSX = "#92D050";
    private final String HEX_STRING_YELLOW_XLSX = "#FFFF00";
    private final String COLOR_CODE_RED = "#FF0000";
    private final String COLOR_CODE_BLUE = "#00B0F0";
    private final String COLOR_CODE_GREEN = "#008000";
    private final String COLOR_CODE_YELLOW = "#FFFF00";
    private final String COLOR_CODE_DARK_GREEN = "#228B22";
    private final String BLUE = "B";
    private final String GREEN = "G";
    private final String YELLOW = "Y";
    private final String RED = "R";
    private RichInputFile inputFileBinding;
    private String message;
    private String downloadMessage;
    private String messageType;
    private String downloadMessageType;
    private String inlineStyle;
    private RichMessage messageBinding;
    private RichMessage downloadMessageBinding;
    private String fileFormat;
    private String fileName;
    private String fileContentType;
    private UploadedFile uploadedFile;
    private RichPanelStretchLayout panelStretchBinding;
    private RichPanelFormLayout inputForm;
    boolean visible = false;
    private TreeSet<EODExcel> eodExcelMainSet = new TreeSet<EODExcel>();
    private TreeSet<EODExcel> eodExcelDeletedSet = new TreeSet<EODExcel>();

    public ExcelBean() {
    }

    public RichTable getExcelTable() {
        return excelTable;
    }

    public void setExcelTable(RichTable xlTab) {
        this.excelTable = xlTab;
    }

    public BindingContainer getContainer() {
        return BindingContext.getCurrent().getCurrentBindingsEntry();
    }

    public void setInputFileBinding(RichInputFile inputFileBinding) {
        this.inputFileBinding = inputFileBinding;
    }

    public RichInputFile getInputFileBinding() {
        return inputFileBinding;
    }

    public void setMessage(String showMessage) {
        this.message = showMessage;
    }

    public String getMessage() {
        return message;
    }

    public void setDownloadMessage(String showMessage) {
        this.downloadMessage = showMessage;
    }

    public String getDownloadMessage() {
        return downloadMessage;
    }

    public void setMessageType(String messageType) {
        this.messageType = messageType;
    }

    public String getMessageType() {
        return messageType;
    }

    public void setDownloadMessageType(String downloadMessageType) {
        this.downloadMessageType = downloadMessageType;
    }

    public String getDownloadMessageType() {
        return downloadMessageType;
    }

    public void setMessageBinding(RichMessage messageBinding) {
        this.messageBinding = messageBinding;
    }

    public RichMessage getMessageBinding() {
        return messageBinding;
    }

    public void setDownloadMessageBinding(RichMessage downloadMessageBinding) {
        this.downloadMessageBinding = downloadMessageBinding;
    }

    public RichMessage getDownloadMessageBinding() {
        return downloadMessageBinding;
    }

    public void setFileFormat(String fileFormat) {
        this.fileFormat = fileFormat;
    }

    public String getFileFormat() {
        return fileFormat;
    }

    public void setFormatMap(HashMap<String, String> formatMap) {
        this.formatMap = formatMap;
    }

    public HashMap<String, String> getFormatMap() {
        return formatMap;
    }

    public void setExtensionMap(HashMap<String, String> extensionMap) {
        this.extensionMap = extensionMap;
    }

    public HashMap<String, String> getExtensionMap() {
        return extensionMap;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileContentType(String fileContentType) {
        this.fileContentType = fileContentType;
    }

    public String getFileContentType() {
        return fileContentType;
    }

    public void setUploadedFile(UploadedFile uploadedFile) {
        this.uploadedFile = uploadedFile;
    }

    public UploadedFile getUploadedFile() {
        return uploadedFile;
    }

    public void setInlineStyle(String inlineStyle) {
        this.inlineStyle = inlineStyle;
    }

    public String getInlineStyle() {
        return inlineStyle;
    }

    public void setPanelStretchBinding(RichPanelStretchLayout panelStretchBinding) {
        this.panelStretchBinding = panelStretchBinding;
    }

    public RichPanelStretchLayout getPanelStretchBinding() {
        return panelStretchBinding;
    }

    public void setInputForm(RichPanelFormLayout inputForm) {
        this.inputForm = inputForm;
    }

    public RichPanelFormLayout getInputForm() {
        return inputForm;
    }

    public void setVisible(boolean visible) {
        this.visible = visible;
    }

    public boolean isVisible() {
        return visible;
    }

    public ViewObject getViewObject(String iteratorName) {
        return ((DCBindingContainer)getContainer()).findIteratorBinding(iteratorName).getViewObject();
    }

    public void executeOperationBinding(String binding) {
        getContainer().getOperationBinding(binding).execute();
    }

    public void cleanupAllRows(ViewObject vo) {
        for (oracle.jbo.Row row : vo.getAllRowsInRange()) {
            row.remove();
        }
    }

    public void initializeFormats() {
        HashMap<String, String> formats = new HashMap<String, String>();
        formats.put("XLSX_FORMAT_1", XLSX_FORMAT_1);
        formats.put("XLSX_FORMAT_2", XLSX_FORMAT_2);
        formats.put("XLS_FORMAT", XLS_FORMAT);
        setFormatMap(formats);
    }

    public void initializeExtensions() {
        HashMap<String, String> extensions = new HashMap<String, String>();
        extensions.put("XLS", XLS);
        extensions.put("XLSX", XLSX);
        setExtensionMap(extensions);
    }

    public void initializeMessages() {
        setMessageType(RichMessage.MESSAGE_TYPE_NONE);
        setMessage("");
    }

    /**
     *
     * @param msg
     * @param msgType
     */
    public void printMessage(String msg, String msgType, String msgColor, RichMessage msgBinding) {
        msgBinding.setInlineStyle(msgBinding.getInlineStyle() + "color:" + msgColor + ";");
        msgBinding.setMessage(msg);
        msgBinding.setMessageType(msgType);
        refreshUIComponent(msgBinding);
    }

    public void refreshUIComponent(UIComponent UIComp) {
        if (UIComp != null) {
            AdfFacesContext.getCurrentInstance().addPartialTarget(UIComp);
        }
    }

    public String getColor(Color color) {
        String colorValue = "";
        if (color instanceof XSSFColor) {
            colorValue = "#" + ((XSSFColor)color).getARGBHex().substring(2);
            if ((colorValue.equalsIgnoreCase(HEX_STRING_BLUE_XLSX)))
                colorValue = BLUE;
            else if ((colorValue.equalsIgnoreCase(HEX_STRING_GREEN_XLSX)))
                colorValue = GREEN;
            else if ((colorValue.equalsIgnoreCase(HEX_STRING_YELLOW_XLSX)))
                colorValue = YELLOW;
            else if ((colorValue.equalsIgnoreCase(COLOR_CODE_RED)))
                colorValue = RED;
            else
                colorValue = null;
        } else if (color instanceof HSSFColor) {
            if (((HSSFColor)color).getHexString().equalsIgnoreCase(HEX_STRING_BLUE_XLS))
                colorValue = COLOR_CODE_BLUE;
            else if (((HSSFColor)color).getHexString().equalsIgnoreCase(HEX_STRING_GREEN_XLS))
                colorValue = COLOR_CODE_GREEN;
            else if (((HSSFColor)color).getHexString().equalsIgnoreCase(HEX_STRING_YELLOW_XLS))
                colorValue = COLOR_CODE_YELLOW;
        }
        return colorValue;
    }

    /**
     *
     * @param valueChangeEvent
     */
    public void uploadExcelVCL(ValueChangeEvent valueChangeEvent) {
        try {
            if (valueChangeEvent != null && valueChangeEvent.getNewValue() != null) {
                initializeFormats();
                initializeExtensions();
                initializeMessages();
                setUploadedFile((UploadedFile)valueChangeEvent.getNewValue());
                setFileName(uploadedFile.getFilename());
                setFileContentType(uploadedFile.getContentType());
                inputFileBinding.setValue(fileName);
                inputFileBinding.setReadOnly(true);
                if (formatMap.containsValue(fileContentType) &&
                    extensionMap.containsValue(fileName.substring(fileName.lastIndexOf(".") + 1,
                                                                  fileName.length()).toUpperCase())) {
                    if (fileName.toUpperCase().endsWith(XLSX) &&
                        (formatMap.get("XLSX_FORMAT_1").equalsIgnoreCase(fileContentType) ||
                         formatMap.get("XLSX_FORMAT_2").equalsIgnoreCase(fileContentType)))
                        setFileFormat(XLSX);
                    else if (fileName.toUpperCase().endsWith(XLS) &&
                             (formatMap.get("XLS_FORMAT").equalsIgnoreCase(fileContentType)))
                        setFileFormat(XLS);
                    else
                        printMessage("Contradiction between format (" +
                                     (fileName.toUpperCase().endsWith(XLSX) ? XLSX : XLS) + ") & content (" +
                                     (fileName.toUpperCase().endsWith(XLSX) ? XLS : XLSX) + ") in the uploaded file",
                                     RichMessage.MESSAGE_TYPE_ERROR, COLOR_CODE_RED, messageBinding);
                } else {
                    printMessage("File format " + fileName.substring(fileName.lastIndexOf("."), fileName.length()) +
                                 " not supported", RichMessage.MESSAGE_TYPE_ERROR, COLOR_CODE_RED, messageBinding);
                }
            } else {
                printMessage("Error in upload", RichMessage.MESSAGE_TYPE_ERROR, COLOR_CODE_RED, messageBinding);
            }
            if (fileFormat != null) {
                readAndLoadXL();
            }
            inputFileBinding.setReadOnly(false);
            printMessage("Excel Processing Completed", RichMessage.MESSAGE_TYPE_CONFIRMATION, COLOR_CODE_GREEN,
                         messageBinding);
            setVisible(true);
        } catch (Exception e) {
            printMessage("Bad File/Format/Inputs!!! Almost everything in the uploaded file is bad!",
                         RichMessage.MESSAGE_TYPE_FATAL, COLOR_CODE_RED, messageBinding);
            e.printStackTrace();
        } finally {
            try {
                inputFileBinding.setReadOnly(false);
                uploadedFile.getInputStream().close();
            } catch (Exception ex) {
                printMessage("Unable to close the connection stream!", RichMessage.MESSAGE_TYPE_FATAL, COLOR_CODE_RED,
                             messageBinding);
                ex.printStackTrace();
            }

        }
    }

    public void readAndLoadXL() throws Exception {
        Workbook workBook =
            fileFormat.equalsIgnoreCase(XLSX) ? new XSSFWorkbook(uploadedFile.getInputStream()) : new HSSFWorkbook(uploadedFile.getInputStream());
        Sheet workSheet = workBook.getSheetAt(0);
        oracle.jbo.Row bcRow = null;
        int iter = 0;
        Cell cell = null;
        ViewObject excelVO = getViewObject(EXCEL_VIEW_ITERATOR);
        cleanupAllRows(excelVO);
        StringBuilder attribute = new StringBuilder();
        for (Row row : workSheet) {
            attribute.delete(0, attribute.length());
            iter = 0;
            executeOperationBinding("CreateInsert");
            bcRow = excelVO.getCurrentRow();
            while (iter < excelVO.getAttributeCount() - 1) {
                if (excelVO.getAttributeDef(iter).getUpdateableFlag() == AttributeDef.UPDATEABLE) {
                    attribute.delete(0, attribute.length());
                    cell = row.getCell(iter);
                    if (cell != null) {
                        if (iter == 4) {
                            Color color = cell.getCellStyle().getFillForegroundColorColor();
                            if (color != null) {
                                bcRow.setAttribute("Color", getColor(color));
                            }
                        }
                        switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            attribute.append(cell.getBooleanCellValue());
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            attribute.append(cell.getErrorCellValue());
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            attribute.append(cell.getCellFormula());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");
                                attribute.append(dateFormat.format(cell.getDateCellValue()));
                            } else {
                                attribute.append(cell.getNumericCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:
                            if (excelVO.getAttributeDef(iter).getName().equalsIgnoreCase("Col9")) {
                                String dateValue = null;
                                SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");
                                try {
                                    dateValue =
                                            cell.getStringCellValue().substring(0, cell.getStringCellValue().indexOf(" ")).replace("-",
                                                                                                                                   "/");
                                    attribute.append(dateFormat.format(new SimpleDateFormat("dd/MM/yyyy").parse(dateValue)));
                                } catch (Exception e) {
                                    try {
                                        attribute.append(dateFormat.format(new SimpleDateFormat("MM/dd/yyyy").parse(dateValue)));
                                    } catch (Exception exo) {
                                        System.out.println("Error in date" + dateValue);
                                    }
                                }
                            } else {
                                attribute.append(cell.getStringCellValue());
                            }
                            break;
                        }
                        bcRow.setAttribute(excelVO.getAttributeDef(iter).getName(),
                                           attribute != null ? String.valueOf(attribute).trim() : null);
                    } else {
                        bcRow.setAttribute("SortOrder", iter);
                    }
                    iter++;
                    System.out.println(attribute != null ? String.valueOf(attribute).trim() : null);
                }
            }
            System.out.println("New Row");
        }
        workBook.close();
    }

    public boolean processRows() throws Exception {
        int count = 0;
        String col1 = null, col5 = null, col6 = null, col7 = null, col8 = null, col9 = null, col14 = null, color =
            null;
        boolean srInd = false, srTypeInd = false, descriptionInd = false, mosStatusInd = false, createdInd =
            false, notesInd = false;
        String sr = null, srType = null, description = null, mosStatus = null, created = null, notes = null;
        EODExcel eodExcelObj = null, temp = null;
        eodExcelMainSet.clear();
        if (eodExcelMainSet.isEmpty()) {
            ViewObject excelVO = getViewObject(EXCEL_VIEW_ITERATOR);
            if (excelVO != null) {
                excelVO.setSortBy("SortOrder asc");
                excelVO.executeQuery();
                oracle.jbo.Row[] allRows = excelVO.getAllRowsInRange();
                for (oracle.jbo.Row oneRow : allRows) {
                    if (nullChecker(oneRow)) {
                        System.out.println("Col1-Color = " + col1 + " - " + color);
                        col1 = (String)oneRow.getAttribute("Col1");
                        color = (String)oneRow.getAttribute("Color");
                        if ((col1 != null && (col1.contains("#") || col1.contains("SR"))) ||
                            (color != null && color.contains(RED))) {
                            continue;
                        } else {
                            if ((col1 != null && !(col1.equalsIgnoreCase("")) &&
                                 col1.replace(".", "").trim().matches("(.*)[0-9](.*)"))) {
                                if (eodExcelObj != null) {
                                    eodExcelMainSet.add(eodExcelObj);
                                    count++;
                                }
                                if (srInd && descriptionInd && createdInd && notesInd) {
                                    temp = eodExcelObj;
                                    eodExcelObj = new EODExcel();
                                    eodExcelObj.setCustomer(temp.getCustomer());
                                    eodExcelObj.setCountry(temp.getCountry());
                                    eodExcelObj.setVersion(temp.getVersion());
                                    eodExcelObj.setSrNo(sr == null ? temp.getSrNo() : sr);
                                    eodExcelObj.setSrType(srType != null ? srType : temp.getSrType());
                                    eodExcelObj.setDescription(description != null ? description :
                                                               temp.getDescription());
                                    eodExcelObj.setDateCreated(created != null ? created : temp.getDateCreated());
                                    eodExcelObj.setNotes(notes != null ? notes : temp.getNotes());
                                    eodExcelObj.setMosStatus(mosStatus != null ? mosStatus : temp.getMosStatus());
                                    eodExcelMainSet.add(eodExcelObj);
                                    count++;
                                    srInd = false;
                                    srTypeInd = false;
                                    descriptionInd = false;
                                    mosStatusInd = false;
                                    createdInd = false;
                                    notesInd = false;
                                    temp = null;
                                    sr = null;
                                    srType = null;
                                    description = null;
                                    mosStatus = null;
                                    created = null;
                                    notes = null;
                                }
                                eodExcelObj = null;
                                eodExcelObj = new EODExcel();
                            } else {
                                if (!srInd && !descriptionInd && !createdInd && !notesInd) {
                                    col5 = (String)oneRow.getAttribute("Col5");
                                    if (col5 != null && !(col5.equalsIgnoreCase("")) &&
                                        (col5.startsWith("--") && (col5.endsWith("--")))) {
                                        oneRow.setAttribute("Col5", null);
                                        srInd = true;
                                    }
                                    col6 = (String)oneRow.getAttribute("Col6");
                                    if (col6 != null && !(col6.equalsIgnoreCase("")) &&
                                        (col6.startsWith("--") && (col6.endsWith("--")))) {
                                        oneRow.setAttribute("Col6", null);
                                        srTypeInd = true;
                                    }
                                    col7 = (String)oneRow.getAttribute("Col7");
                                    if (col7 != null && !(col7.equalsIgnoreCase("")) &&
                                        (col7.startsWith("--") && (col7.endsWith("--")))) {
                                        oneRow.setAttribute("Col7", null);
                                        descriptionInd = true;
                                    }
                                    col8 = (String)oneRow.getAttribute("Col8");
                                    if (col8 != null && !(col8.equalsIgnoreCase("")) &&
                                        (col8.startsWith("--") && (col8.endsWith("--")))) {
                                        oneRow.setAttribute("Col8", null);
                                        mosStatusInd = true;
                                    }
                                    col9 = (String)oneRow.getAttribute("Col9"); //Created
                                    if (col9 != null && !(col9.equalsIgnoreCase("")) &&
                                        (col9.startsWith("--") && (col9.endsWith("--")))) {
                                        oneRow.setAttribute("Col9", null);
                                        createdInd = true;
                                    }
                                    col14 = (String)oneRow.getAttribute("Col14");
                                    if (col14 != null && !(col14.equalsIgnoreCase("")) &&
                                        (col14.startsWith("--") && (col14.endsWith("--")))) {
                                        oneRow.setAttribute("Col14", null);
                                        notesInd = true;
                                    }
                                }
                            }
                        }
                        if (eodExcelObj != null) {
                            if (createdInd) {
                                created =
                                        (created != null && !(created.equalsIgnoreCase(""))) ? created.trim() + "\r\n" +
                                        (oneRow.getAttribute("Col9") != null &&
                                         !(oneRow.getAttribute("Col9").equals("")) ?
                                         ((String)oneRow.getAttribute("Col9")).trim() : "") :
                                        oneRow.getAttribute("Col9") != null &&
                                        !(oneRow.getAttribute("Col9").equals("")) ?
                                        ((String)oneRow.getAttribute("Col9")).trim() : "";
                            } else {
                                if (oneRow.getAttribute("Col9") != null &&
                                    !(oneRow.getAttribute("Col9").equals(""))) //Date Created
                                    eodExcelObj.setDateCreated(((String)oneRow.getAttribute("Col9")).trim());
                            }
                            if (srInd) {
                                sr = (sr != null && !(sr.equalsIgnoreCase(""))) ? sr.trim() + "\r\n" +
                                        (oneRow.getAttribute("Col5") != null &&
                                         !(oneRow.getAttribute("Col5").equals("")) ?
                                         ((String)oneRow.getAttribute("Col5")).trim() : "") :
                                        oneRow.getAttribute("Col5") != null &&
                                        !(oneRow.getAttribute("Col5").equals("")) ?
                                        ((String)oneRow.getAttribute("Col5")).trim() : "";
                            } else {
                                if (oneRow.getAttribute("Col5") != null &&
                                    !(oneRow.getAttribute("Col5").equals(""))) //SR#
                                    eodExcelObj.setSrNo(((String)oneRow.getAttribute("Col5")).trim());
                            }
                            if (descriptionInd) {
                                description =
                                        (description != null && !(description.equalsIgnoreCase(""))) ? description.trim() +
                                        "\r\n" +
                                        (oneRow.getAttribute("Col7") != null &&
                                         !(oneRow.getAttribute("Col7").equals("")) ?
                                         ((String)oneRow.getAttribute("Col7")).trim() : "") :
                                        oneRow.getAttribute("Col7") != null &&
                                        !(oneRow.getAttribute("Col7").equals("")) ?
                                        ((String)oneRow.getAttribute("Col7")).trim() : "";
                            } else {
                                if (eodExcelObj.getDescription() != null &&
                                    !(eodExcelObj.getDescription().equalsIgnoreCase(""))) //Description
                                    eodExcelObj.setDescription(eodExcelObj.getDescription() != null ?
                                                               eodExcelObj.getDescription().trim() :
                                                               eodExcelObj.getDescription() + "\r\n" +
                                            (((String)oneRow.getAttribute("Col7")).trim()));
                                else if (oneRow.getAttribute("Col7") != null &&
                                         !(oneRow.getAttribute("Col7").equals("")))
                                    eodExcelObj.setDescription(((String)oneRow.getAttribute("Col7")).trim());
                            }
                            if (srTypeInd) {
                                srType = (srType != null && !(srType.equalsIgnoreCase(""))) ? srType.trim() + "\r\n" +
                                        (oneRow.getAttribute("Col6") != null &&
                                         !(oneRow.getAttribute("Col6").equals("")) ?
                                         ((String)oneRow.getAttribute("Col6")).trim() : "") :
                                        oneRow.getAttribute("Col6") != null &&
                                        !(oneRow.getAttribute("Col6").equals("")) ?
                                        ((String)oneRow.getAttribute("Col6")).trim() : "";
                            } else {
                                if (oneRow.getAttribute("Col6") != null &&
                                    !(oneRow.getAttribute("Col6").equals(""))) //SR Type
                                    eodExcelObj.setSrType(((String)oneRow.getAttribute("Col6")).trim());
                            }
                            if (oneRow.getAttribute("Col2") != null &&
                                !(oneRow.getAttribute("Col2").equals(""))) //Customer
                                eodExcelObj.setCustomer(((String)oneRow.getAttribute("Col2")).trim());
                            if (oneRow.getAttribute("Col4") != null && !(oneRow.getAttribute("Col4").equals("")) &&
                                //Version
                                eodExcelObj.getVersion() == null)
                                eodExcelObj.setVersion(((String)oneRow.getAttribute("Col4")).trim());
                            if (oneRow.getAttribute("Col4") != null && !(oneRow.getAttribute("Col4").equals("")) &&
                                eodExcelObj.getVersion() != null) //Country
                                eodExcelObj.setCountry(((String)oneRow.getAttribute("Col4")).trim());
                            if (mosStatusInd) {
                                mosStatus =
                                        (mosStatus != null && !(mosStatus.equalsIgnoreCase(""))) ? mosStatus.trim() +
                                        "\r\n" +
                                        (oneRow.getAttribute("Col8") != null &&
                                         !(oneRow.getAttribute("Col8").equals("")) ?
                                         ((String)oneRow.getAttribute("Col8")).trim() : "") :
                                        oneRow.getAttribute("Col8") != null &&
                                        !(oneRow.getAttribute("Col8").equals("")) ?
                                        ((String)oneRow.getAttribute("Col8")).trim() : "";

                            } else {
                                if (oneRow.getAttribute("Col8") != null &&
                                    !(oneRow.getAttribute("Col8").equals(""))) //MOS Status
                                    eodExcelObj.setMosStatus(((String)oneRow.getAttribute("Col8")).trim());
                            }
                            if (notesInd) {
                                notes = (notes != null && !(notes.equalsIgnoreCase(""))) ? notes.trim() + "\r\n" +
                                        (oneRow.getAttribute("Col14") != null &&
                                         !(oneRow.getAttribute("Col14").equals("")) ?
                                         ((String)oneRow.getAttribute("Col14")).trim() : "") :
                                        oneRow.getAttribute("Col14") != null &&
                                        !(oneRow.getAttribute("Col14").equals("")) ?
                                        ((String)oneRow.getAttribute("Col14")).trim() : "";
                            } else {
                                if (eodExcelObj.getNotes() != null &&
                                    !(eodExcelObj.getNotes().equalsIgnoreCase(""))) //Notes
                                    eodExcelObj.setNotes(eodExcelObj.getNotes() + "\r\n" +
                                            (((String)oneRow.getAttribute("Col14"))).trim());
                                else if (oneRow.getAttribute("Col14") != null &&
                                         !(oneRow.getAttribute("Col14").equals("")))
                                    eodExcelObj.setNotes(((String)oneRow.getAttribute("Col14")).trim());
                            }
                            if (eodExcelObj.getColorOrder() == null && oneRow.getAttribute("Color") != null &&
                                !(oneRow.getAttribute("Color").equals(""))) //Color
                                eodExcelObj.setColorOrder(((String)oneRow.getAttribute("Color")).equalsIgnoreCase("G") ?
                                                          "1" :
                                                          (((String)oneRow.getAttribute("Color")).equalsIgnoreCase("Y") ?
                                                           "2" : "3"));
                        }
                    }
                }
                if (eodExcelObj != null)
                    eodExcelMainSet.add(eodExcelObj);
            }
        }
        return true;
    }

    public boolean printExcel(java.io.OutputStream outputStream) throws Exception {
        long maxString = 0;
        Workbook workbook = new XSSFWorkbook();
        CreationHelper helper = workbook.getCreationHelper();
        Sheet sheet1 = workbook.createSheet("EOD Stabilization Tracker");
        Sheet sheet2 = null;
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short)12);
        headerFont.setColor(IndexedColors.RED.getIndex());
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        Row mainHeaderRow = sheet1.createRow(0);
        Cell mainHeadCell = mainHeaderRow.createCell(0);
        mainHeadCell.setCellValue("Main Entries");
        mainHeadCell.setCellStyle(headerStyle);
        Row headerRow = sheet1.createRow(1);
        String[] columns =
        { "Date Created", "SR#", "Description", "SR Type", "Customer", "Version", "Country", "MOS Status", "Notes" };
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerStyle);
        }
        Font bodyFont = workbook.createFont();
        bodyFont.setFontHeightInPoints((short)10);
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(helper.createDataFormat().getFormat("dd-MMM-yy"));
        dateCellStyle.setFont(bodyFont);
        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.setFont(bodyFont);
        CellStyle bigCellStyle = workbook.createCellStyle();
        bigCellStyle.setFont(bodyFont);
        bigCellStyle.setWrapText(true);
        int rowNum = 2;
        if (eodExcelMainSet != null && eodExcelMainSet.size() > 0) {
            Iterator<EODExcel> eodExcelIterator = eodExcelMainSet.iterator();
            while (eodExcelIterator.hasNext()) {
                EODExcel eodEx = eodExcelIterator.next();
                if (eodEx != null) {
                    Row row = sheet1.createRow(rowNum++);
                    if (eodEx.getNotes() != null && eodEx.getDescription() != null)
                        maxString =
                                eodEx.getNotes().length() > eodEx.getDescription().length() ? eodEx.getNotes().length() :
                                eodEx.getDescription().length();

                    row.setHeight((short)(((maxString / 40) < 4 ? 4 :
                                           ((maxString / 40) > 10 ? 10 : (maxString / 40))) *
                                          sheet1.getDefaultRowHeight()));

                    Cell dateCreatedCell = row.createCell(0);
                    dateCreatedCell.setCellValue(eodEx.getDateCreated());
                    dateCreatedCell.setCellStyle(dateCellStyle);

                    Cell srNoCell = row.createCell(1);
                    srNoCell.setCellValue(eodEx.getSrNo());
                    srNoCell.setCellStyle(bodyStyle);

                    Cell descriptionCell = row.createCell(2);
                    descriptionCell.setCellValue(eodEx.getDescription());
                    descriptionCell.setCellStyle(bigCellStyle);

                    Cell srTypeCell = row.createCell(3);
                    srTypeCell.setCellValue(eodEx.getSrType());
                    srTypeCell.setCellStyle(bodyStyle);

                    Cell customerCell = row.createCell(4);
                    customerCell.setCellValue(eodEx.getCustomer());
                    customerCell.setCellStyle(bodyStyle);

                    Cell versionCell = row.createCell(5);
                    versionCell.setCellValue(eodEx.getVersion());
                    versionCell.setCellStyle(bodyStyle);

                    Cell countryCell = row.createCell(6);
                    countryCell.setCellValue(eodEx.getCountry());
                    countryCell.setCellStyle(bodyStyle);

                    Cell mosStatusCell = row.createCell(7);
                    mosStatusCell.setCellValue(eodEx.getMosStatus());
                    mosStatusCell.setCellStyle(bodyStyle);

                    Cell notesCell = row.createCell(8);
                    notesCell.setCellValue(eodEx.getNotes());
                    notesCell.setCellStyle(bigCellStyle);
                }
            }
        }
        maxString = 0;
        rowNum = 0;
        if (eodExcelDeletedSet != null && eodExcelDeletedSet.size() > 0) {
            sheet2 = workbook.createSheet("Duplicate Entries");
            Row deletedHead = sheet2.createRow(rowNum++);
            Cell deletedHeadCell = deletedHead.createCell(0);
            deletedHeadCell.setCellValue("Duplicate Entries");
            deletedHeadCell.setCellStyle(headerStyle);
            Row deletedHeaderRow = sheet2.createRow(rowNum++);
            for (int i = 0; i < columns.length; i++) {
                Cell deletedCell = deletedHeaderRow.createCell(i);
                deletedCell.setCellValue(columns[i]);
                deletedCell.setCellStyle(headerStyle);
            }

            Iterator<EODExcel> eodExcelDeletedIterator = eodExcelDeletedSet.iterator();
            while (eodExcelDeletedIterator.hasNext()) {
                EODExcel eodExDeleted = eodExcelDeletedIterator.next();
                if (eodExDeleted != null) {
                    Row row = sheet2.createRow(rowNum++);
                    if (eodExDeleted.getNotes() != null && eodExDeleted.getDescription() != null)
                        maxString =
                                eodExDeleted.getNotes().length() > eodExDeleted.getDescription().length() ? eodExDeleted.getNotes().length() :
                                eodExDeleted.getDescription().length();

                    row.setHeight((short)(((maxString / 40) < 4 ? 4 :
                                           ((maxString / 40) > 10 ? 10 : (maxString / 40))) *
                                          sheet2.getDefaultRowHeight()));

                    Cell dateCreatedCell = row.createCell(0);
                    dateCreatedCell.setCellValue(eodExDeleted.getDateCreated());
                    dateCreatedCell.setCellStyle(dateCellStyle);

                    Cell srNoCell = row.createCell(1);
                    srNoCell.setCellValue(eodExDeleted.getSrNo());
                    srNoCell.setCellStyle(bodyStyle);

                    Cell descriptionCell = row.createCell(2);
                    descriptionCell.setCellValue(eodExDeleted.getDescription());
                    descriptionCell.setCellStyle(bigCellStyle);

                    Cell srTypeCell = row.createCell(3);
                    srTypeCell.setCellValue(eodExDeleted.getSrType());
                    srTypeCell.setCellStyle(bodyStyle);

                    Cell customerCell = row.createCell(4);
                    customerCell.setCellValue(eodExDeleted.getCustomer());
                    customerCell.setCellStyle(bodyStyle);

                    Cell versionCell = row.createCell(5);
                    versionCell.setCellValue(eodExDeleted.getVersion());
                    versionCell.setCellStyle(bodyStyle);

                    Cell countryCell = row.createCell(6);
                    countryCell.setCellValue(eodExDeleted.getCountry());
                    countryCell.setCellStyle(bodyStyle);

                    Cell mosStatusCell = row.createCell(7);
                    mosStatusCell.setCellValue(eodExDeleted.getMosStatus());
                    mosStatusCell.setCellStyle(bodyStyle);

                    Cell notesCell = row.createCell(8);
                    notesCell.setCellValue(eodExDeleted.getNotes());
                    notesCell.setCellStyle(bigCellStyle);
                }
            }
        }
        for (int i = 0; i < columns.length; i++) {
            if (i != 2 && i != 8) {
                sheet1.autoSizeColumn(i);
                if (sheet2 != null)
                    sheet2.autoSizeColumn(i);
            } else {
                sheet1.setColumnWidth(i, 15000);
                if (sheet2 != null)
                    sheet2.setColumnWidth(i, 15000);
            }
        }
        sheet1.createFreezePane(0, 2);
        if (sheet2 != null)
            sheet2.createFreezePane(0, 2);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        workbook.close();
        return true;
    }

    public boolean clearDuplicates() throws Exception {
        ArrayList<EODExcel> toBeRemovedList = new ArrayList<EODExcel>();
        EODExcel currSet = null, prevSet = null;
        Iterator<EODExcel> eodExcelIter = eodExcelMainSet.iterator();
        while (eodExcelIter.hasNext()) {
            currSet = eodExcelIter.next();
            if (prevSet != null && prevSet.getSrNo() != null && currSet.getSrNo() != null &&
                prevSet.getSrNo().equalsIgnoreCase(currSet.getSrNo())) {
                if (currSet != null && currSet.getColorOrder() != null &&
                    !(currSet.getColorOrder().equalsIgnoreCase("1"))) {
                    if ((Integer.parseInt(prevSet.getColorOrder()) < Integer.parseInt(currSet.getColorOrder()))) {
                        eodExcelDeletedSet.add(currSet);
                        eodExcelIter.remove();
                    } else if ((Integer.parseInt(prevSet.getColorOrder()) ==
                                Integer.parseInt(currSet.getColorOrder()))) {
                        if (prevSet.getNotes().length() > currSet.getNotes().length()) {
                            eodExcelDeletedSet.add(currSet);
                            eodExcelIter.remove();
                        } else {
                            eodExcelDeletedSet.add(prevSet);
                            toBeRemovedList.add(prevSet);
                        }
                    }
                } else
                    prevSet = currSet;
            } else {
                prevSet = currSet;
            }
        }
        if (toBeRemovedList != null && toBeRemovedList.size() > 0) {
            for (int i = 0; i < toBeRemovedList.size(); i++) {
                eodExcelMainSet.remove(toBeRemovedList.get(i));
            }
        }
        return true;
    }

    public void downloadExcel(javax.faces.context.FacesContext facesContext, java.io.OutputStream outputStream) {
        try {
            if (facesContext != null) {
                if (processRows()) {
                    if (clearDuplicates()) {
                        if (printExcel(outputStream))
                            printMessage("Excel Generation Completed!", RichMessage.MESSAGE_TYPE_CONFIRMATION,
                                         COLOR_CODE_DARK_GREEN, downloadMessageBinding);
                    }
                } else {
                    printMessage("Error in Downloading Excel", RichMessage.MESSAGE_TYPE_ERROR, COLOR_CODE_RED,
                                 downloadMessageBinding);
                }
            }
        } catch (Exception e) {
            printMessage("Error in Downloading Excel", RichMessage.MESSAGE_TYPE_ERROR, COLOR_CODE_RED,
                         downloadMessageBinding);
            e.printStackTrace();
        }
    }

    boolean nullChecker(oracle.jbo.Row row) {
        boolean result = false;
        for (int i = 0; i < getViewObject(EXCEL_VIEW_ITERATOR).getAttributeCount(); i++) {
            if (row.getAttribute(i) != null && !(row.getAttribute(i).equals(""))) {
                result = true;
            }
        }
        return result;
    }


    //    public static void main(String[] args) {
    //        String a = "7.0";
    //        a = a.replace(".", "").trim();
    //        System.out.println(a+" - "+a.replace(".", ""));
    //        System.out.println(a.matches("[^0-9]"));
    //        System.out.println(a.matches("(.*)\\d(.*)"));
    //        System.out.println(a.matches("[0-9]"));
    //        System.out.println(a.matches("(.*)^[0-9](.*)"));
    //        System.out.println(a.replace(".", "").matches("[^0-9]"));
    //        System.out.println(a.replace(".", "").matches("\\d"));
    //        a="-";
    //        System.out.println(a.replace(".", "").matches("(.*)[0-9](.*)"));
    //        System.out.println(a.replace(".", "").matches("^[0-9]"));
    //    }

}
