package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;

import java.io.File;
import java.util.List;


public class SendToDCCInit extends UnifiedAgent {
    ISession session;
    IDocumentServer server;
    IBpmService bpm;
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    IUser user;
    IUser owner;
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
        user = session.getUser();
        task = getEventTask();

        try {

            helper = new ProcessHelper(session);
            (new File(Conf.SendToDCC.MainPath)).mkdirs();

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


            String ownCode = "";
            if(Utils.hasDescriptor(projectInfObj, Conf.Descriptors.ProjectOwn)){
                ownCode = projectInfObj.getDescriptorValue(Conf.Descriptors.ProjectOwn, String.class);
                ownCode = (ownCode == null ? "" : ownCode);
            }
            if(ownCode.isEmpty()){
                throw new Exception("Project owner is empty.");
            }
            IInformationObject pown = Utils.getContractor(ownCode, helper);
            if(pown == null){
                throw new Exception("Project-owner not found [" + ownCode + "].");
            }

            sendToDCCLinks = processInstance.getLoadedInformationObjectLinks();
            Utils.verifyProcessSubDocuments(sendToDCCLinks, projectNo);

            processInstance = Utils.updateProcessInstance(processInstance);

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

            String ownName = "";
            if(Utils.hasDescriptor(pown, Conf.Descriptors.ContractorName)){
                ownName = pown.getDescriptorValue(Conf.Descriptors.ContractorName, String.class);
                ownName = (ownName == null ? "" : ownName);
            }

            processInstance.setDescriptorValue(Conf.Descriptors.ReceiverCode, ownCode);
            processInstance.setDescriptorValue(Conf.Descriptors.ReceiverName, ownName);

            processInstance.setDescriptorValue(Conf.Descriptors.SenderCode, supCode);
            processInstance.setDescriptorValue(Conf.Descriptors.SenderName, supName);

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