package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.ser.blueline.bpm.IWorkbasket;
import com.ser.blueline.metaDataComponents.IStringMatrix;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.core.spreadsheet.HTMLOptions;

import jakarta.activation.DataHandler;
import jakarta.activation.DataSource;
import jakarta.activation.FileDataSource;
import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.UUID;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Utils {

    public static void loadExcel(String tpth, String name, JSONObject pbks) throws IOException {

        FileInputStream tist = new FileInputStream(tpth);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);

        Sheet tsht = twrb.getSheet(name);
        for (Row trow : tsht){
            for(Cell tcll : trow){
                if(tcll.getCellType() != CellType.STRING){continue;}
                String clvl = tcll.getRichStringCellValue().getString();
                String clvv = updateCell(clvl, pbks);
                if(!clvv.equals(clvl)){
                    tcll.setCellValue(clvv);
                }

                if(clvv.indexOf("[[") != (-1) && clvv.indexOf("]]") != (-1)
                        && clvv.indexOf("[[") < clvv.indexOf("]]")){
                    String znam = clvv.substring(clvv.indexOf("[[") + "[[".length(), clvv.indexOf("]]"));
                    if(pbks.has(znam)){
                        String zval = znam;
                        if(pbks.has(znam + ".Text")){
                            zval = pbks.getString(znam + ".Text");
                        }
                        tcll.setCellValue(zval);
                        String lurl = pbks.getString(znam);
                        if(!lurl.isEmpty()) {
                            Hyperlink link = twrb.getCreationHelper().createHyperlink(HyperlinkType.URL);
                            link.setAddress(lurl);
                            tcll.setHyperlink(link);
                        }
                    }
                }
            }
        }
        FileOutputStream tost = new FileOutputStream(tpth);
        twrb.write(tost);
        tost.close();
    }
    static void sendResultMail(IBpmService bpm,
                               ISession session, IDocumentServer server,
                               ITask task, IInformationObject project, String prjn, String akey, String nots,
                               JSONObject mailConfig,
                               IInformationObjectLinks links,
                               ProcessHelper helper) throws Exception {

        if(task.getProcessInstance() == null){return;}
        IUser ownr = task.getProcessInstance().getCreator();
        String tMail = ownr.getEMailAddress();
        tMail = (tMail == null ? "" : tMail);
        if(tMail.isEmpty()){return;}

        IDocument ptpl = getMailTplDocument(prjn, helper);
        if(ptpl == null){return;}

        System.out.println("  ---> " + ptpl.getDisplayName());
        JSONObject ecfg = getExcelConfig(ptpl, prjn);
        if(ecfg == null){return;}

        String uniqueId = UUID.randomUUID().toString();
        String mailExcelPath = Utils.exportDocument(ptpl, Conf.SendToDCC.MainPath, "[" + prjn + "]@" + akey + "@[" + uniqueId + "]");

        Double dcix = (ecfg.has("Document.Lines.ColumnIndex") ? ecfg.getDouble("Document.Lines.ColumnIndex") : 0.0d);
        loadTableRows(mailExcelPath, ecfg.getString("SheetName"), "Document",
                (dcix == null ? 0 : (int) Math.round(dcix)), links.getLinks().size());

        String[] cc = getPrjMails(bpm, project, ecfg, akey + ".Mail-CC");

        JSONObject mbms = new JSONObject();

        mbms.put("DoxisLink", mailConfig.getString("webBase") + helper.getTaskURL(task.getID()));

        mbms.put("Count", links.getLinks().size() + "");
        mbms.put("Result", akey);
        mbms.put("Comment", nots);

        int dcnt = 0;
        for (ILink link : links.getLinks()) {
            IDocument xdoc = (IDocument) link.getTargetInformationObject();
            if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}


            String docNo = "", revNo = "", docName = "";
            if(hasDescriptor(xdoc, Conf.Descriptors.DocNumber)){
                docNo = xdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                docNo = (docNo == null ? "" : docNo);
            }
            if(hasDescriptor(xdoc, Conf.Descriptors.DocRevision)){
                revNo = xdoc.getDescriptorValue(Conf.Descriptors.DocRevision, String.class);
                revNo = (revNo == null ? "" : revNo);
            }
            if(hasDescriptor(xdoc, Conf.Descriptors.DocName)){
                docName = xdoc.getDescriptorValue(Conf.Descriptors.DocName, String.class);
                docName = (docName == null ? "" : docName);
            }

            dcnt++;
            mbms.put("DocNo" + dcnt, docNo);
            mbms.put("RevNo" + dcnt, revNo);
            mbms.put("DocumentName" + dcnt, docName);
        }
        loadExcel(mailExcelPath, ecfg.getString("SheetName"), mbms);

        String mailHtmlPath = convertExcelToHtml(mailExcelPath,
                Conf.SendToDCC.MainPath + "/[" + prjn + "]@" + akey + "@[" + uniqueId + "]" + ".html", "");
        JSONObject mail = new JSONObject();

        mail.put("To", tMail);
        mail.put("CC", String.join(";", cc));
        mail.put("Subject",
                "Send To DCC Result {ProjectNo} / {Result}"
                        .replace("{ProjectNo}", prjn)
                        .replace("{Result}", akey)
        );
        mail.put("BodyHTMLFile", mailHtmlPath);

        try {
            Utils.sendHTMLMail(session, server, mailConfig, mail);
        } catch (Exception ex){
            System.out.println("EXCP [Send-Mail] : " + ex.getMessage());
        }
    }
    private static IDocument getMailTplDocument(String prjn, ProcessHelper helper){
        IDocument dtpl = Utils.getTemplateDocument(prjn, Conf.SendToDCC.MailTemplate, helper);
        return dtpl;
    }

    private static Row copyRow(org.apache.poi.ss.usermodel.Workbook workbook, Sheet worksheet, int sourceRowNum, int destinationRowNum) {

        Row newRow = worksheet.getRow(destinationRowNum);
        Row sourceRow = worksheet.getRow(sourceRowNum);

        if (newRow != null) {
            worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
        } else {
            newRow = worksheet.createRow(destinationRowNum);
        }

        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {

            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);

            if (oldCell == null) {
                continue;
            }


            CellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);


            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }


            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }


            newCell.setCellType(oldCell.getCellType());


            switch (oldCell.getCellType()) {
                case BLANK:// Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
                default:
                    break;
            }
        }

        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
                        cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }

        return newRow;
    }
    public static Row getMasterRow(Sheet sheet, String prfx, Integer colIx)  {
        for (Row row : sheet) {
            Cell cll1 = row.getCell(colIx);
            if(cll1 == null){continue;}

            String cval = cll1.getRichStringCellValue().getString();
            if(cval.isEmpty()){continue;}

            if(!cval.equals("[*" + prfx + "*]") ){continue;}
            return row;

        }
        return null;
    }
    public static void loadTableRows(String spth, String shtName, String prfx, Integer colIx, Integer scpy) throws IOException {

        FileInputStream tist = new FileInputStream(spth);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);

        Sheet tsht = twrb.getSheet(shtName);
        Row mrow = getMasterRow(tsht, prfx, colIx);
        if(mrow == null){return ;}

        mrow.getCell(colIx).setBlank();

        for(var i=1;i<=scpy;i++){
            Row nrow = copyRow(twrb, tsht, mrow.getRowNum(), mrow.getRowNum() + i);

            for(Cell ncll : nrow) {
                if (ncll.getCellType() != CellType.STRING) {
                    continue;
                }
                if(ncll.getColumnIndex() == colIx){
                    ncll.setBlank();
                    continue;
                }

                String clvl = ncll.getRichStringCellValue().getString();
                String clvv = clvl.replace("*", i+"");
                if(!clvv.equals(clvl)){
                    ncll.setCellValue(clvv);
                }
            }
        }

        mrow.setZeroHeight(true);
        tsht.setColumnHidden(colIx, true);

        FileOutputStream tost = new FileOutputStream(spth);
        twrb.write(tost);
        tost.close();

    }

    private static String[] getPrjMails(IBpmService bpm, IInformationObject project, JSONObject ecfg, String znam){
        List<String> rtrn = new ArrayList<>();
        List<String> list = new ArrayList<>();
        List<Object> cvls = (ecfg.has(znam) ? (JSONArray) ecfg.get(znam) : new JSONArray()).toList();
        for(Object cval : cvls){
            String sval = (String) cval;
            if(sval.isEmpty()){continue;}

            if(sval.equals("Project Mngr.")){
                String pmng = project.getDescriptorValue(Conf.Descriptors.ProjectMngr, String.class);
                if(pmng != null && !pmng.isEmpty() && !list.contains(pmng)){
                    list.add(pmng);
                }
            }
            if(sval.equals("Engineering Manager")){
                String emng = project.getDescriptorValue(Conf.Descriptors.EngMngr, String.class);
                if(emng != null && !emng.isEmpty() && !list.contains(emng)){
                    list.add(emng);
                }
            }
            if(sval.equals("DCC List")){
                List<String> dlst = project.getDescriptorValues(Conf.Descriptors.DCCList, String.class);
                for(String dccu : dlst){
                    if(dccu != null && !dccu.isEmpty() && !list.contains(dccu)){
                        list.add(dccu);
                    }
                }
            }
        }

        for(String line : list){
            IWorkbasket lwbk = bpm.getWorkbasket(line);
            if(lwbk == null){continue;}

            String wbMail = lwbk.getNotifyEMail();
            if(wbMail == null || wbMail.isEmpty()){continue;}

            if(rtrn.contains(wbMail)){continue;}
            rtrn.add(wbMail);
        }

        return rtrn.toArray(new String[rtrn.size()]);
    }
    private static JSONObject getExcelConfig(IDocument template, String prjn) throws Exception {
        String excelPath = FileEvents.fileExport(template, Conf.SendToDCC.MainPath, "[" + prjn + "]");
        JSONObject ecfg = (FilenameUtils.getExtension(excelPath).toString().toUpperCase().equals("XLSX") ?
                Utils.getXlsxConfig(excelPath) : new JSONObject());
        if(!ecfg.has("SheetName")){
            return null;
        }
        return ecfg;
    }
    public static JSONObject getXlsxConfig(String excelPath) throws Exception {
        JSONObject rtrn = new JSONObject();

        FileInputStream fist = new FileInputStream(excelPath);
        XSSFWorkbook fwrb = new XSSFWorkbook(fist);
        Sheet sheet = fwrb.getSheet("#CONFIG");
        if(sheet == null){throw new Exception("#CONFIG sheet not found. (" + excelPath + ")");}

        for(Row row : sheet) {
            Cell cll1 = row.getCell(0);
            if(cll1 == null){continue;}

            Cell cll3 = row.getCell(2);
            if(cll3 == null){continue;}

            if(cll1.getCellType() != CellType.STRING){continue;}
            String cnam = cll1.getStringCellValue().trim();
            if(cnam.isEmpty()){continue;}

            String ctyp = "String";
            Cell cll2 = row.getCell(1);
            if(cll2 != null) {
                CellType ttyp = cll2.getCellType();
                if (ttyp == CellType.STRING) {
                    ctyp = cll2.getStringCellValue().trim();
                }
            }

            CellType tval = cll3.getCellType();
            if(tval == CellType.STRING && ctyp.equals("String")) {
                String cvalString = cll3.getStringCellValue().trim();
                rtrn.put(cnam, cvalString);
            }
            if(tval == CellType.NUMERIC && ctyp.equals("Numeric")) {
                Double cvalNumeric = cll3.getNumericCellValue();
                rtrn.put(cnam, cvalNumeric);
            }
            if(tval == CellType.STRING && ctyp.equals("List")) {
                String cvalString = cll3.getStringCellValue().trim();
                List<Object> cvls = (rtrn.has(cnam) ? (JSONArray) rtrn.get(cnam) : new JSONArray()).toList();
                if(cvalString != null && !cvalString.isEmpty() && !cvls.contains(cvalString)) {
                    cvls.add(cvalString);
                    rtrn.put(cnam, cvls);
                }
            }
        }
        return rtrn;
    }
    static IInformationObject getContractor(String scod, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.Supplier).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.ObjectNumber).append(" = '").append(scod).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.SupplierContact} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return informationObjects[0];
    }
    static IInformationObject getContact(String mail, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.Contact).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrimaryEMail).append(" = '").append(mail).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.SupplierContact} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return informationObjects[0];
    }
    static void verifyProcessSubDocuments(IInformationObjectLinks links, String prjNo) throws Exception{

        List<String> docIds = new ArrayList<>();
        JSONObject rmvs = new JSONObject();

        for (ILink link : links.getLinks()) {
            IDocument xdoc = (IDocument) link.getTargetInformationObject();
            if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
            String xdId = xdoc.getID();
            if (docIds.contains(xdId)){continue;}

            String dpjn = xdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            dpjn = (dpjn == null ? "" : dpjn);

            String dsts = xdoc.getDescriptorValue(Conf.Descriptors.DocStatus, String.class);
            dsts = (dsts == null ? "" : dsts);

            if(!Conf.CheckValues.SendDocStatuses.contains(dsts)
            || !dpjn.equals(prjNo)){
                if(!rmvs.has(xdId)){
                    rmvs.put(xdId, xdoc);
                }
                continue;
            }

            docIds.add(xdoc.getID());
        }
        for(String rmId : rmvs.keySet()){
            IDocument rdoc = (IDocument) rmvs.get(rmId);
            System.out.println("Remove documents : " + rmId);
            links.removeInformationObject(rmId, false);
        }
    }
    static void updateProcessSubDocuments(ISession ses, ProcessHelper helper, IInformationObjectLinks links, String prjNo, String status, String notes) throws Exception{
        List<String> docIds = new ArrayList<>();
        for (ILink link : links.getLinks()) {
            IDocument xdoc = (IDocument) link.getTargetInformationObject();
            if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
            String xdId = xdoc.getID();
            if (docIds.contains(xdId)){continue;}

            String dpjn = xdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            dpjn = (dpjn == null ? "" : dpjn);

            String dsts = xdoc.getDescriptorValue(Conf.Descriptors.DocStatus, String.class);
            dsts = (dsts == null ? "" : dsts);

            if(!Conf.CheckValues.SendDocStatuses.contains(dsts)
            || !dpjn.equals(prjNo)){
                continue;
            }
            docIds.add(xdoc.getID());
        }
        for(String docId : docIds){
            IInformationObject pdoc = ses.getDocumentServer().getInformationObjectByID(docId, ses);
            if(pdoc == null){continue;}
            if(!hasDescriptor(pdoc, Conf.Descriptors.DocStatus)){continue;}
            pdoc.setDescriptorValue(Conf.Descriptors.DocStatus, status);

            if(hasDescriptor(pdoc, Conf.Descriptors.Notes) && !notes.isEmpty()){
                pdoc.setDescriptorValue(Conf.Descriptors.Notes, notes);
            }

            String docNo = "", revNo = "";
            if(hasDescriptor(pdoc, Conf.Descriptors.DocNumber)){
                docNo = pdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                docNo = (docNo == null ? "" : docNo);
            }
            if(hasDescriptor(pdoc, Conf.Descriptors.DocRevision)){
                revNo = pdoc.getDescriptorValue(Conf.Descriptors.DocRevision, String.class);
                revNo = (revNo == null ? "" : revNo);
            }

            if(helper != null && !docNo.isEmpty() && !revNo.isEmpty()){
                //updateDocReleased(docNo, revNo, helper);
                pdoc.setDescriptorValue(Conf.Descriptors.Released, "1");
            }
            updateInfoObj(pdoc);
        }
    }
    static JSONObject getSystemConfig(ISession ses) throws Exception {
        return getSystemConfig(ses, null);
    }
    static IProcessInstance updateProcessInstance(IProcessInstance prin) throws Exception {
        String prInId = prin.getID();
        prin.commit();
        Thread.sleep(2000);
        if(prInId.equals("<new>")) {
            return prin;
        }
        return (IProcessInstance) prin.getSession().getDocumentServer().getInformationObjectByID(prInId, prin.getSession());
    }
    static IInformationObject updateInfoObj(IInformationObject info) throws Exception {
        String prInId = info.getID();
        info.commit();
        Thread.sleep(2000);
        if(prInId.equals("<new>")) {
            return info;
        }
        return (IInformationObject) info.getSession().getDocumentServer().getInformationObjectByID(prInId, info.getSession());
    }
    static IInformationObject getProjectWorkspace(String prjn, ProcessHelper helper) {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.ProjectWorkspace).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjCardCode).append(" = '").append(prjn).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.ProjectWorkspace} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return informationObjects[0];
    }
    public static boolean hasDescriptor(IInformationObject infObj, String dscn) throws Exception {
        IValueDescriptor[] vds = infObj.getDescriptorList();
        for(IValueDescriptor vd : vds){
            if(vd.getName().equals(dscn)){return true;}
        }
        return false;
    }
    static String projectNr(IInformationObject projectInfObj) throws Exception {
        String rtrn = "";
        if(Utils.hasDescriptor(projectInfObj, Conf.Descriptors.ProjectNo)){
            rtrn = projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            rtrn = (rtrn == null ? "" : rtrn).trim();
        }
        return rtrn;
    }
    static JSONObject
    getSystemConfig(ISession ses, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = ses.getDocumentServer().getStringMatrix("CCM_SYSTEM_CONFIG", ses);
        }
        if(mtrx == null) throw new Exception("SystemConfig Global Value List not found");

        List<List<String>> rawTable = mtrx.getRawRows();

        String srvn = ses.getSystem().getName().toUpperCase();
        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            String name = line.get(0);
            if(!name.toUpperCase().startsWith(srvn + ".")){continue;}
            name = name.substring(srvn.length() + ".".length());
            rtrn.put(name, line.get(1));
        }
        return rtrn;
    }
    static String updateCell(String str, JSONObject bookmarks){
        StringBuffer rtr1 = new StringBuffer();
        String tmp = str + "";
        Pattern ptr1 = Pattern.compile( "\\{([\\w\\.]+)\\}" );
        Matcher mtc1 = ptr1.matcher(tmp);
        while(mtc1.find()) {
            String mk = mtc1.group(1);
            String mv = "";
            if(bookmarks.has(mk)){
                mv = bookmarks.getString(mk);
            }
            mtc1.appendReplacement(rtr1,  mv);
        }
        mtc1.appendTail(rtr1);
        tmp = rtr1.toString();

        return tmp;
    }
    static String exportDocument(IDocument document, String exportPath, String fileName) throws IOException {
        String rtrn ="";
        IDocumentPart partDocument = document.getPartDocument(document.getDefaultRepresentation() , 0);
        String fName = (!fileName.isEmpty() ? fileName : partDocument.getFilename());
        fName = fName.replaceAll("[\\\\/:*?\"<>|]", "_");
        try (InputStream inputStream = partDocument.getRawDataAsStream()) {
            IFDE fde = partDocument.getFDE();
            if (fde.getFDEType() == IFDE.FILE) {
                rtrn = exportPath + "/" + fName + "." + ((IFileFDE) fde).getShortFormatDescription();

                try (FileOutputStream fileOutputStream = new FileOutputStream(rtrn)){
                    byte[] bytes = new byte[2048];
                    int length;
                    while ((length = inputStream.read(bytes)) > -1) {
                        fileOutputStream.write(bytes, 0, length);
                    }
                }
            }
        }
        return rtrn;
    }
    static String saveDocReviewExcel(String templatePath, Integer shtIx, String tpltSavePath, JSONObject pbks) throws IOException {

        FileInputStream tist = new FileInputStream(templatePath);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);

        Sheet tsht = twrb.getSheetAt(shtIx);
        for (Row trow : tsht){
            for(Cell tcll : trow){
                if(tcll.getCellType() != CellType.STRING){continue;}
                String clvl = tcll.getRichStringCellValue().getString();
                String clvv = updateCell(clvl, pbks);
                if(!clvv.equals(clvl)){
                    tcll.setCellValue(clvv);
                }

                if(clvv.indexOf("[[") != (-1) && clvv.indexOf("]]") != (-1)
                        && clvv.indexOf("[[") < clvv.indexOf("]]")){
                    String znam = clvv.substring(clvv.indexOf("[[") + "[[".length(), clvv.indexOf("]]"));
                    if(pbks.has(znam)){
                        tcll.setCellValue(znam);
                        String lurl = pbks.getString(znam);
                        if(!lurl.isEmpty()) {
                            Hyperlink link = twrb.getCreationHelper().createHyperlink(HyperlinkType.URL);
                            link.setAddress(lurl);
                            tcll.setHyperlink(link);
                        }
                    }
                }
            }
        }
        FileOutputStream tost = new FileOutputStream(tpltSavePath);
        twrb.write(tost);
        tost.close();
        return tpltSavePath;
    }
    static IDocument getTemplateDocument(String prjNo, String tpltName, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.Template).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjCardCode).append(" = '").append(prjNo).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.ObjectNumberExternal).append(" = '").append(tpltName).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.Company} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    static String convertExcelToHtml(String excelPath, String htmlPath, String s)  {
        Workbook workbook = new Workbook();
        workbook.loadFromFile(excelPath);
        Worksheet sheet = workbook.getWorksheets().get(0);
        HTMLOptions options = new HTMLOptions();
        options.setImageEmbedded(true);
        sheet.saveToHtml(htmlPath, options);
        return htmlPath;
    }
    static String getFileContent (String path) throws Exception {
        return new String(Files.readAllBytes(Paths.get(path)));
    }
    static String
    getHTMLFileContent (String path) throws Exception {
        String rtrn = new String(Files.readAllBytes(Paths.get(path)));
        rtrn = rtrn.replace("\uFEFF", "");
        rtrn = rtrn.replace("ï»¿", "");
        return rtrn;
    }
    static IStringMatrix getMailConfigMatrix(ISession ses, IDocumentServer srv, String mtpn) throws Exception {
        IStringMatrix rtrn = srv.getStringMatrix("CCM_MAIL_CONFIG", ses);
        if (rtrn == null) throw new Exception("MailConfig Global Value List not found");
        return rtrn;
    }
    static JSONObject getMailConfig(ISession ses, IDocumentServer srv, String mtpn) throws Exception {
        return getMailConfig(ses, srv, mtpn, null);
    }
    static JSONObject getMailConfig(ISession ses, IDocumentServer srv, String mtpn, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = getMailConfigMatrix(ses, srv, mtpn);
        }
        if(mtrx == null) throw new Exception("MailConfig Global Value List not found");
        List<List<String>> rawTable = mtrx.getRawRows();

        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            rtrn.put(line.get(0), line.get(1));
        }
        return rtrn;
    }
    static void sendHTMLMail(ISession ses, IDocumentServer srv, JSONObject mcfg, JSONObject pars) throws Exception {
        String host = mcfg.getString("host");
        String port = mcfg.getString("port");
        String protocol = mcfg.getString("protocol");
        String sender = mcfg.getString("sender");
        String subject = "";
        String mailTo = "";
        String mailCC = "";
        String attachments = "";

        if(pars.has("From")){
            sender = pars.getString("From");
        }
        if(pars.has("To")){
            mailTo = pars.getString("To");
        }

        if(sender.isEmpty()){throw new Exception("Mail Sender is empty");}
        if(mailTo.isEmpty()){throw new Exception("Mail To is empty");}

        if(pars.has("CC")){
            mailCC = pars.getString("CC");
        }
        if(pars.has("Subject")){
            subject = pars.getString("Subject");
        }
        if(pars.has("AttachmentPaths")){
            attachments = pars.getString("AttachmentPaths");
        }


        Properties props = new Properties();
        props.put("mail.debug","true");
        props.put("mail.smtp.debug", "true");

        props.put("mail.smtp.host", host);
        props.put("mail.smtp.port", port);

        String start_tls = (mcfg.has("start_tls") ? mcfg.getString("start_tls") : "");
        if(start_tls.equals("true")) {
            props.put("mail.smtp.starttls.enable", start_tls);
        }

        String auth = mcfg.getString("auth");
        props.put("mail.smtp.auth", auth);
        jakarta.mail.Authenticator authenticator = null;
        if(!auth.equals("false")) {
            String auth_username = mcfg.getString("auth.username");
            String auth_password = mcfg.getString("auth.password");

            if (host.contains("gmail")) {
                props.put("mail.smtp.socketFactory.port", port);
                props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
                props.put("mail.smtp.socketFactory.fallback", "false");
            }
            if (protocol != null && protocol.contains("TLSv1.2"))  {
                props.put("mail.smtp.ssl.protocols", protocol);
                props.put("mail.smtp.ssl.trust", "*");
                props.put("mail.smtp.socketFactory.port", port);
                props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
                props.put("mail.smtp.socketFactory.fallback", "false");
            }
            authenticator = new jakarta.mail.Authenticator(){
                @Override
                protected jakarta.mail.PasswordAuthentication getPasswordAuthentication(){
                    return new jakarta.mail.PasswordAuthentication(auth_username, auth_password);
                }
            };
        }
        props.put("mail.mime.charset","UTF-8");
        Session session = (authenticator == null ? Session.getDefaultInstance(props) : Session.getDefaultInstance(props, authenticator));

        MimeMessage message = new MimeMessage(session);
        message.setFrom(new InternetAddress(sender.replace(";", ",")));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(mailTo.replace(";", ",")));
        message.setRecipients(Message.RecipientType.CC, InternetAddress.parse(mailCC.replace(";", ",")));
        message.setSubject(subject);

        Multipart multipart = new MimeMultipart("mixed");

        BodyPart htmlBodyPart = new MimeBodyPart();
        htmlBodyPart.setContent(getHTMLFileContent(pars.getString("BodyHTMLFile")) , "text/html; charset=UTF-8"); //5
        multipart.addBodyPart(htmlBodyPart);

        String[] atchs = attachments.split("\\;");
        for (String atch : atchs){
            if(atch.isEmpty()){continue;}
            BodyPart attachmentBodyPart = new MimeBodyPart();
            attachmentBodyPart.setDataHandler(new DataHandler((DataSource) new FileDataSource(atch)));

            String fnam = Paths.get(atch).getFileName().toString();
            if(pars.has("AttachmentName." + fnam)){
                fnam = pars.getString("AttachmentName." + fnam);
            }

            attachmentBodyPart.setFileName(fnam);
            multipart.addBodyPart(attachmentBodyPart);

        }

        message.setContent(multipart);
        Transport.send(message);

    }

}
