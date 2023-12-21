package ser;

import org.json.JSONObject;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Conf {

    public static class MainWFUpdateSheetIndex {
        public static final Integer Mail = 0;
    }
    public static class Databases{
        public static final String Company = "D_QCON";
        public static final String EngineeringDocument= "PRJ_DOC";
        public static final String ProjectWorkspace = "PRJ_FOLDER";
    }
    public static class SendToDCC {
        public static final String MainPath = "C:/tmp2/bulk/send-to-dcc";
    }
    public static class ClassIDs{
        public static final String Template = "b9cf43d1-a4d3-482f-9806-44ae64c6139d";
        public static final String EngineeringDocument = "3b3078f8-c0d0-4830-b963-88bebe1c1462";
        public static final String ProjectWorkspace = "32e74338-d268-484d-99b0-f90187240549";
    }
    public static class Descriptors{
        public static final String ProjectNo = "ccmPRJCard_code";
        public static final String DocNumber = "ccmPrjDocNumber";
        public static final String Revision = "ccmPrjDocRevision";
        public static final String DocStatus = "ccmPrjDocStatus";
        public static final List<String> DocStatuses = new ArrayList<>(Arrays.asList(
            "10",
            "20"
        ));

    }
    public static class DescriptorLiterals{
        public static final String PrjCardCode = "CCMPRJCARD_CODE";
        public static final String ObjectNumberExternal = "OBJECTNUMBER2";
    }
}
