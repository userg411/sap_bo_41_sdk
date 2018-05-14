import com.crystaldecisions.sdk.occa.infostore.*; 
import com.crystaldecisions.sdk.plugin.desktop.folder.*; 
import com.crystaldecisions.sdk.exception.SDKException;
import com.crystaldecisions.sdk.framework.CrystalEnterprise;
import com.crystaldecisions.sdk.framework.IEnterpriseSession;
import com.crystaldecisions.sdk.framework.ISessionMgr;
import com.crystaldecisions.sdk.occa.infostore.IInfoObject;
import com.crystaldecisions.sdk.occa.infostore.IInfoObjects;
import com.crystaldecisions.sdk.occa.infostore.IInfoStore;
import com.crystaldecisions.sdk.plugin.desktop.user.IUser;
import com.crystaldecisions.sdk.plugin.desktop.usergroup.IUserGroup;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Font;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.util.*; 
import org.apache.poi.ss.usermodel.CellStyle;


public class GetSec{ 
   // Set the following three variables with logon info the the CMS. 
    public static final String strUser = ""; 
    public static final String strCMS = ""; 
    public static final String strPwd = ""; 
    
   // Set this to true to display all advanced rights; false will just display a count of assigned rights if the access level is "advanced" 
    public static Boolean showAdvancedDetail = false; 
    public static IInfoStore infoStore; 
   
    public static PrintWriter out;
    public static String strDelimeter;
   
   
    public static HSSFWorkbook workbook;
    public static HSSFSheet sheet;
    public static CellStyle boldStyle;
    public static Font font;
           
    public static int counter = 0;
   
   
    public static void main(String[] args) throws SDKException, FileNotFoundException , Exception{ 
        strDelimeter = "~";
    
        // Log in to CMS and get infoStore 
        ISessionMgr oSessionMgr; 
        IEnterpriseSession oEnterpriseSession; 
        IInfoObjects iObjects;
      
        oSessionMgr = CrystalEnterprise.getSessionMgr(); 
        oEnterpriseSession = oSessionMgr.logon(strUser, strPwd, strCMS,  "secEnterprise"); 
        infoStore = (IInfoStore)oEnterpriseSession.getService("", "InfoStore"); 

        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet("Sample sheet");
        
        boldStyle = workbook.createCellStyle();//Create style
        Font font = workbook.createFont();//Create font
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold
        boldStyle.setFont(font);//set it to bold
        
        
        traverseFolders(id, infoStore,"");
       
        
        FileOutputStream out = new FileOutputStream(new File(""));
        workbook.write(out);
        out.close();
        System.out.println("Excel written successfully..");
      
   }
   private static void printGroupMembers(IInfoStore infoStore) throws SDKException{
        IInfoObjects infoObjects = (IInfoObjects)infoStore.query("select * from ci_systemobjects where si_kind = 'Usergroup'");
        Class[] interfaces = infoObjects.get(0).getClass().getInterfaces();
        
        for(int i=0; i<infoObjects.size();i++){
            IUserGroup group  = (IUserGroup)infoObjects.get(i);
            System.out.println("Group name: "+group.getTitle());
            Set <Integer> set = group.getUsers();
            
            for(int j:set){
               IInfoObject userObject = (IInfoObject)infoStore.query("select * from ci_systemobjects where si_id = "+j).get(0);
               IUser user = (IUser)userObject;
               System.out.println("\t"+user.getTitle());
            }
        }
    }
   static void printUsers(IInfoStore oInfoStore) throws SDKException{
        IInfoObjects users = oInfoStore.query("Select * From CI_SYSTEMOBJECTS  Where SI_KIND='USER'");
       
        for(int i=0;i<users.size();i++)
        {
            IInfoObject userObject = (IInfoObject)users.get(i);
            IUser user = (IUser)userObject;
            Set <Integer> set = user.getGroups();
            for(int j:set){
                IInfoObject groupObject = (IInfoObject)infoStore.query("select * from ci_systemobjects where si_id = "+j).get(0);
                IUserGroup userGroup = (IUserGroup)groupObject;
                System.out.println("\t"+userGroup.getTitle());
            }
        }
       
       
   }
   
   static void traverseFolders(int id,IInfoStore oInfoStore, String delimiter) throws SDKException, Exception
   {
        IInfoObject infoObject = (IInfoObject)oInfoStore.query("Select SI_ID From CI_INFOOBJECTS " + "Where SI_id="+id).get(0);
       
        if("Webi".equals(infoObject.getKind()))
        {
           //System.out.println(delimiter+"Report: "+infoObject.getTitle());
           return;
        }
        else if ("Folder".equals(infoObject.getKind()))
        {
            //System.out.println(delimiter+"Folder: "+infoObject.getTitle());
            listReportRights(oInfoStore, id,delimiter, counter);
            counter++;
            IInfoObjects infoObjects = oInfoStore.query("Select SI_ID From CI_INFOOBJECTS Where SI_parentid = "+id);
            for(int i=0;i<infoObjects.size();i++)
            {
                IInfoObject report = (IInfoObject)infoObjects.get(i);
                traverseFolders(report.getID(),oInfoStore, delimiter+"    ");
            }
       }
       else return;
   }
   static void traverseUniverses(int id,IInfoStore oInfoStore, String delimiter) throws SDKException, Exception
   {
        IInfoObject infoObject = (IInfoObject)oInfoStore.query("Select SI_ID From CI_APPOBJECTS " + "Where SI_id="+id).get(0);
       
        if("Universe".equals(infoObject.getKind()))
        {
           System.out.println(delimiter+"Universe: "+infoObject.getTitle());
           listReportRights(oInfoStore, id,delimiter, counter);
           return;
        }
        else if ("Folder".equals(infoObject.getKind()))
        {
            System.out.println(delimiter+"Folder: "+infoObject.getTitle());
            listReportRights(oInfoStore, id,delimiter, counter);
            counter++;
            IInfoObjects infoObjects = oInfoStore.query("Select SI_ID From CI_APPOBJECTS Where SI_parentid = "+id);
            for(int i=0;i<infoObjects.size();i++)
            {
                IInfoObject report = (IInfoObject)infoObjects.get(i);
                traverseUniverses(report.getID(),oInfoStore, delimiter+"    ");
            }
       }
       else return;
   }
   static void traverseConnections(int id,IInfoStore oInfoStore, String delimiter) throws Exception{
        IInfoObjects infoObjects = oInfoStore.query("Select SI_ID From CI_APPOBJECTS Where si_kind = 'MetaData.DataConnection'");
        for(int i=0;i<infoObjects.size();i++)
            {
                IInfoObject report = (IInfoObject)infoObjects.get(i);
                listReportRights(oInfoStore, report.getID(),delimiter, counter);
                counter++;
                
            }
       
   }
   static void listReportRights(IInfoStore oInfoStore,int folderId, String delimiter, int rowCount) throws SDKException,Exception{
       
       
       IInfoObjects infoObjects = oInfoStore.query("Select SI_ID From CI_INFOOBJECTS " + "Where SI_id = " + folderId);
        IInfoObject report = (IInfoObject)infoObjects.get(0);
        Row row = sheet.createRow(counter);
        
        int col = delimiter.length()/4;
        
        Cell cell = row.createCell(col);
        cell.setCellStyle(boldStyle);
        cell.setCellValue(report.getKind()+": "+report.getTitle());
        counter++;
       
        System.out.println(delimiter+"Folder: "+report.getTitle());
        
        ISecurityInfo2 securityInfo = report.getSecurityInfo2();
        IEffectivePrincipals effectivePrincipals = securityInfo.getEffectivePrincipals();
        
        
        
        Iterator it = effectivePrincipals.iterator();
        while (it.hasNext()){
            IEffectivePrincipal effectivePrincipal = (IEffectivePrincipal)it.next();
            IEffectiveRoles effectiveRoles = effectivePrincipal.getRoles();
            Iterator roleIT = effectiveRoles.iterator();
            Row row1 = sheet.createRow(counter);
            counter++;
            Cell cell1 = row1.createCell(col);
            System.out.print(delimiter+effectivePrincipal.getName() +" has rights: ");
            String rights = "";
            
            while (roleIT.hasNext()){
                IEffectiveRole effectiveRole = (IEffectiveRole)roleIT.next();
                rights+=effectiveRole.getTitle()+"; ";
                System.out.print(effectiveRole.getTitle()+"; ");
            }
            cell1.setCellValue(effectivePrincipal.getName() + " has rights: " + rights );
            IEffectiveRights effectiveRights = effectivePrincipal.getRights();
            Iterator rightIT = effectiveRights.iterator();
            
        System.out.println();
        }
    }
  
    static Integer printEm(IInfoStore oInfoStore,IInfoObjects iObjects) throws SDKException{ 
        Integer maxID = new Integer(99999999); 
        for(int i = 0; i < iObjects.size(); i++){ 
            IInfoObject iObject = (IInfoObject) iObjects.get(i); 
            maxID = new Integer(iObject.getID()); 
            String outString;
            outString = iObject.getID() + strDelimeter + iObject.getKind() + strDelimeter + getObjectPath(iObject); 
            out.println(outString);
        } 
      return maxID; 
    } 
   // Get the full path of an object 
    static String getObjectPath(IInfoObject inObject) throws SDKException{ 
        IInfoObject oIO = inObject; 
        String path = ""; 

        while(true){ 
             // If the current object is a folder, get its "si_path" info; otherwise just iterate up through the objects' parents 
            if ("Folder".equals(oIO.getKind())) 
            { 
                oIO = (IInfoObject) infoStore.query("select si_id,si_path from CI_infoobjects,ci_systemobjects,ci_appobjects where si_id = " + oIO.getID()).get(0); 
                path = oIO.getTitle() +  path; 
                try { 
                    if (oIO.getParentID() != 0) 
                        for(String pathPart : ((IFolder)oIO).getPath() ) 
                            path = pathPart + "/" + path; 
                } 
                catch (Exception wtf) 
                { 
                    // This shouldn't happen since we're checking for path count of 0 above, but just in case... 
                    return "<" + oIO.getID() + ">" + "/" + path + oIO.getTitle(); 
                } 
                return path; 
             } 
            else 
                path = path + "/" + oIO.getTitle(); 
            oIO = oIO.getParent(); 
        } 
   } 
    
}
