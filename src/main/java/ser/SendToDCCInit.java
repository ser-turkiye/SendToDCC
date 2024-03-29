package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.util.List;


public class SendToDCCInit extends UnifiedAgent {
    Logger log = LogManager.getLogger();
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    IUser owner;
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

            processInstance = task.getProcessInstance();
            owner = processInstance.getOwner();

            projectNo = (processInstance != null ? Utils.projectNr((IInformationObject) processInstance) : "");
            if(projectNo.isEmpty()){
                throw new Exception("Project no is empty.");
            }

            //sender & receiver code+name set

            projectInfObj = Utils.getProjectWorkspace(projectNo, helper);
            if(projectInfObj == null){
                throw new Exception("Project not found [" + projectNo + "].");
            }


            String ownCode = Utils.getMainCompGVList(projectNo);
            String ownName = Utils.getMainCompNameGVList(projectNo);

            if(ownCode.isEmpty()){
                throw new Exception("Project owner is empty.");
            }


            sendToDCCLinks = processInstance.getLoadedInformationObjectLinks();
            Utils.verifyProcessSubDocuments(sendToDCCLinks, projectNo);

            String status = "40", draft = "10";

            Utils.updateProcessSubDocuments(sendToDCCLinks, projectNo, draft, status, "", false);
            processInstance.commit();

            IInformationObject cont = Utils.getContact(owner.getLogin(), helper);
            if(cont == null){
                throw new Exception("Contact not found [" + owner.getLogin() + "].");
            }
            String supCode = "";
            if(Utils.hasDescriptor(cont, Conf.Descriptors.ContractorCode)){
                supCode = cont.getDescriptorValue(Conf.Descriptors.ContractorCode, String.class);
                supCode = (supCode == null ? "" : supCode);
            }
            if(supCode.isEmpty()){
                throw new Exception("Supplier code is empty.");
            }


            String supName = "";
            if(Utils.hasDescriptor(cont, Conf.Descriptors.ContractorName)){
                supName = cont.getDescriptorValue(Conf.Descriptors.ContractorName, String.class);
                supName = (supName == null ? "" : supName);
            }


            processInstance.setDescriptorValue(Conf.Descriptors.ReceiverCode, ownCode);
            processInstance.setDescriptorValue(Conf.Descriptors.ReceiverName, ownName);

            processInstance.setDescriptorValue(Conf.Descriptors.SenderCode, supCode);
            processInstance.setDescriptorValue(Conf.Descriptors.SenderName, supName);

            processInstance.commit();
            log.info("Tested.");

        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultError("Exception : " + e.getMessage());

        }

        log.info("Finished");
        return resultSuccess("Ended successfully");
    }
}