package ser;

import com.ser.blueline.IDocumentServer;
import com.ser.blueline.IInformationObject;
import com.ser.blueline.IInformationObjectLinks;
import com.ser.blueline.ISession;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.json.JSONObject;

import java.io.File;
import java.util.List;


public class SendToDCCRejected extends UnifiedAgent {
    ISession session;
    IDocumentServer server;
    IBpmService bpm;
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    IInformationObject contractorInfObj;
    IInformationObjectLinks sendToDCCLinks;
    ProcessHelper helper;
    ITask task;
    List<String> documentIds;
    String transmittalNr;
    String projectNo;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        if(getEventTask().getProcessInstance().findLockInfo().getOwnerID() != null){
            return resultRestart("Restarting Agent");
        }

        session = getSes();
        bpm = getBpm();
        server = session.getDocumentServer();
        task = getEventTask();

        try {

            helper = new ProcessHelper(session);
            (new File(Conf.SendToDCC.MainPath)).mkdirs();

            JSONObject scfg = Utils.getSystemConfig(session);
            if(scfg.has("LICS.SPIRE_XLS")){
                com.spire.license.LicenseProvider.setLicenseKey(scfg.getString("LICS.SPIRE_XLS"));
            }


            processInstance = task.getProcessInstance();
            projectNo = (processInstance != null ? Utils.projectNr((IInformationObject) processInstance) : "");
            if(projectNo.isEmpty()){
                throw new Exception("Project no is empty.");
            }

            //sender & receiver code+name set

            projectInfObj = Utils.getProjectWorkspace(projectNo, helper);
            if(projectInfObj == null){
                throw new Exception("Project not found [" + projectNo + "].");
            }

            String ivpNo = processInstance.getDescriptorValue(Conf.Descriptors.SenderCode, String.class);
            if(ivpNo == null || ivpNo.isEmpty()){
                throw new Exception("Involve Party code is empty.");
            }
            contractorInfObj = Utils.getContractorFolder(projectNo, ivpNo, helper);
            if(contractorInfObj == null){
                throw new Exception("Involve Party [" + projectNo + "/" + ivpNo + "].");
            }

            sendToDCCLinks = processInstance.getLoadedInformationObjectLinks();
            String notes = "";
            if(Utils.hasDescriptor((IInformationObject) processInstance, Conf.Descriptors.Notes)){
                notes = processInstance.getDescriptorValue(Conf.Descriptors.Notes, String.class);
                notes = (notes == null ? "" : notes);
            }
            Utils.updateProcessSubDocuments(session, null, sendToDCCLinks, projectNo, "90", notes);

            JSONObject mcfg = Utils.getMailConfig(session, server, "");
            Utils.sendResultMail(bpm, session, server, task, projectInfObj, projectNo, contractorInfObj, ivpNo, "Rejected", notes, mcfg, sendToDCCLinks, helper);
            processInstance = Utils.updateProcessInstance(processInstance);
            System.out.println("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            System.out.println("Exception       : " + e.getMessage());
            System.out.println("    Class       : " + e.getClass());
            System.out.println("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(), 10);

        }

        System.out.println("Finished");
        return resultSuccess("Ended successfully");
    }
}