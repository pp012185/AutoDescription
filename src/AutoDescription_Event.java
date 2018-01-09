import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.IDataObject;
import com.agile.api.INode;
import com.agile.px.*;

public class AutoDescription_Event implements IEventAction{
    @Override
    public EventActionResult doAction(IAgileSession iAgileSession, INode iNode, IEventInfo req) {

        System.out.println("------Start------");

        try {
            IObjectEventInfo info = (IObjectEventInfo)req;
            // getDataObject()
            IDataObject obj = info.getDataObject();
            String objNumber = obj.getName();
            System.out.println("ObjName: "+objNumber);



        } catch (APIException e) {
            e.printStackTrace();
        }





        return new EventActionResult(req, new ActionResult(0,"Success"));
    }
}
