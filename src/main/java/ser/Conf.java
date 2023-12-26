package ser;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Conf {

    public static class Databases{
        public static final String Company = "D_QCON";
        public static final String EngineeringDocument= "PRJ_DOC";
        public static final String ProjectWorkspace = "PRJ_FOLDER";
        public static final String SupplierContact = "BPWS";
    }
    public static class SendToDCC {
        public static final String MainPath = "C:/tmp2/bulk/send-to-dcc";
        public static final String MailTemplate = "SEND_TO_DCC_RESULT_MAIL";
    }
    public static class ClassIDs{
        public static final String Template = "b9cf43d1-a4d3-482f-9806-44ae64c6139d";
        public static final String EngineeringDocument = "3b3078f8-c0d0-4830-b963-88bebe1c1462";
        public static final String ProjectWorkspace = "32e74338-d268-484d-99b0-f90187240549";
        public static final String Contact = "d7ffea9d-3419-4922-8ffa-a0310add5723";
        public static final String Supplier = "4fd133c1-4cf8-461e-bb09-a39c307feb50";
    }
    public static class Descriptors{
        public static final String ProjectNo = "ccmPRJCard_code";
        public static final String DocNumber = "ccmPrjDocNumber";
        public static final String DocRevision = "ccmPrjDocRevision";
        public static final String DocStatus = "ccmPrjDocStatus";
        public static final String DocName = "ObjectName";
        public static final String Notes = "Notes";
        public static final String ProjectOwn = "ccmPRJCard_prefix";
        public static final String ContractorCode = "ObjectNumber";
        public static final String ContractorName = "ObjectName";
        public static final String SenderName = "ccmTrmtSender";
        public static final String SenderCode = "ccmTrmtSenderCode";
        public static final String ReceiverName = "ccmTrmtReceiver";
        public static final String ReceiverCode = "ccmTrmtReceiverCode";
        public static final String Released = "ccmReleased";

        public static final String ProjectMngr = "ccmPRJCard_prjmngr";
        public static final String EngMngr = "ccmPRJCard_EngMng";
        public static final String DCCList = "ccmPrjCard_DccList";
    }
    public static class CheckValues{
        public static final List<String> SendDocStatuses = new ArrayList<>(Arrays.asList(
            "",
            "10",
            "20",
            "40"
        ));

    }
    public static class DescriptorLiterals{
        public static final String PrjCardCode = "CCMPRJCARD_CODE";
        public static final String ObjectNumberExternal = "OBJECTNUMBER2";
        public static final String PrimaryEMail = "PRIMARYEMAIL";
        public static final String ObjectNumber = "OBJECTNUMBER";
        public static final String DocNumber = "CCMPRJDOCNUMBER";
        public static final String DocRevision = "CCMPRJDOCREVISION";
    }
}
