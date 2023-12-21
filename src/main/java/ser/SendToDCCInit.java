package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;

import java.io.File;
import java.util.ArrayList;
import java.util.List;


public class SendToDCCInit extends UnifiedAgent {
    ISession session;
    IDocumentServer server;
    IBpmService bpm;
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
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


            processInstance = task.getProcessInstance();
            projectNo = (processInstance != null ? Utils.projectNr((IInformationObject) processInstance) : "");
            if(projectNo.isEmpty()){
                throw new Exception("Project no is empty.");
            }

            projectInfObj = Utils.getProjectWorkspace(projectNo, helper);
            if(projectInfObj == null){
                throw new Exception("Project not found [" + projectNo + "].");
            }
            sendToDCCLinks = processInstance.getLoadedInformationObjectLinks();
            documentIds = new ArrayList<>();
            List<String> rmvIds = new ArrayList<>();

            for (ILink link : sendToDCCLinks.getLinks()) {
                IDocument xdoc = (IDocument) link.getTargetInformationObject();
                if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                String xdId = xdoc.getID();
                if (documentIds.contains(xdId)){continue;}

                String dsts = xdoc.getDescriptorValue(Conf.Descriptors.DocStatus, String.class);
                dsts = (dsts == null ? "" : dsts);
                if(!Conf.Descriptors.DocStatuses.contains(dsts)){
                    if(!rmvIds.contains(xdId)){
                        rmvIds.add(xdId);
                    }
                    continue;
                }

                documentIds.add(xdoc.getID());
            }
            for(String rmId : rmvIds){
                sendToDCCLinks.removeInformationObject(rmId, false);
            }

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