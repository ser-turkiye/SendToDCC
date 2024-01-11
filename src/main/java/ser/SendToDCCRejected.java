package ser;

import com.ser.blueline.IDocumentServer;
import com.ser.blueline.IInformationObject;
import com.ser.blueline.IInformationObjectLinks;
import com.ser.blueline.ISession;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.io.File;
import java.util.List;


public class SendToDCCRejected extends UnifiedAgent {
    Logger log = LogManager.getLogger();
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    IInformationObject contractorInfObj;
    IInformationObjectLinks sendToDCCLinks;
    ProcessHelper helper;
    ITask task;
    String projectNo;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        if(getEventTask().getProcessInstance().findLockInfo().getOwnerID() != null){
            return resultRestart("Restarting Agent");
        }

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.SendToDCC.MainPath);

        task = getEventTask();

        try {

            helper = new ProcessHelper(Utils.session);

            JSONObject scfg = Utils.getSystemConfig();
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

            /*
            String ivpNo = processInstance.getDescriptorValue(Conf.Descriptors.SenderCode, String.class);
            if(ivpNo == null || ivpNo.isEmpty()){
                throw new Exception("Involve Party code is empty.");
            }
            contractorInfObj = Utils.getContractorFolder(projectNo, ivpNo, helper);
            if(contractorInfObj == null){
                throw new Exception("Involve Party [" + projectNo + "/" + ivpNo + "].");
            }
            */

            sendToDCCLinks = processInstance.getLoadedInformationObjectLinks();
            String notes = "";
            if(Utils.hasDescriptor((IInformationObject) processInstance, Conf.Descriptors.Notes)){
                notes = processInstance.getDescriptorValue(Conf.Descriptors.Notes, String.class);
                notes = (notes == null ? "" : notes);
            }
            String stsAction = "Rejected";
            String status = "90", waiting = "40";

            Utils.updateProcessSubDocuments(sendToDCCLinks, projectNo, waiting, status, notes, false);
            processInstance.commit();

            JSONObject mcfg = Utils.getMailConfig();
            Utils.sendResultMail(Conf.SendToDCC.MailTemplate, task, projectInfObj, projectNo,
                    //contractorInfObj, ivpNo,
                    stsAction, notes, mcfg, sendToDCCLinks, helper);

            log.info("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(), 10);

        }

        log.info("Finished");
        return resultSuccess("Ended successfully");
    }
}