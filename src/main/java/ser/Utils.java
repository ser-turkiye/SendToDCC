package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
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

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Utils {
    static void updateDocReleased(String dcod, String rcod, ProcessHelper helper) throws Exception {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.EngineeringDocument).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.DocNumber).append(" = '").append(dcod).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.DocRevision).append(" = '").append(rcod).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] infoObjs = helper.createQuery(new String[]{Conf.Databases.EngineeringDocument} , whereClause , 1);
        for(IInformationObject info : infoObjs){
            if(!hasDescriptor(info, Conf.Descriptors.Released)){continue;}
            info.setDescriptorValue(Conf.Descriptors.Released, "0");
            info = updateInfoObj(info);
        }
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

            if(hasDescriptor(pdoc, Conf.Descriptors.Notes)){
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
                updateDocReleased(docNo, revNo, helper);
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
    static String convertExcelToHtml(String excelPath, String htmlPath)  {
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
    static void sendHTMLMail(ISession ses, IDocumentServer srv, String mtpn, JSONObject pars) throws Exception {
        JSONObject mcfg = Utils.getMailConfig(ses, srv, mtpn);

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
