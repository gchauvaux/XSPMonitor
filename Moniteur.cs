using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Web.SessionState;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Text.RegularExpressions;
using System.IO;
using System.CodeDom.Compiler;
using System.Reflection;
//using Microsoft.JScript;
using Microsoft.CSharp;
using System.Data.OracleClient;
using System.Net.Mail;
using ClassLib;


/// <summary>
/// Description résumée de MoniteurNet
/// </summary>

public static class K
{
    public static HttpSessionState Session= HttpContext.Current.Session;
    public static HttpRequest Request= HttpContext.Current.Request;
    public static HttpResponse Response = HttpContext.Current.Response;
    public static HttpServerUtility Server= HttpContext.Current.Server;
    public static string getString(object o)
    {
        if (o == null || o == System.DBNull.Value)
            return null;
        return Convert.ToString(o);
    }
    public static int getInt(object o)
    {
        if (o==null || o == System.DBNull.Value)
            return 0;
        return Convert.ToInt32(o);
    }
    public static Decimal getDecimal(object o)
    {
        if (o == null || o == System.DBNull.Value)
            return 0;
        return Convert.ToDecimal(o);
    }
    public static string defaut(string pdef, string pval)
    {
        if (pval == null || pval=="")
            return pdef;
        else
            return pval;
    }
    public static object getObject(object o)
    {
       // if (o == System.DBNull.Value)
       //     return "";
        return o;
    }
    public static string rowclass(int i)
    {
        return "clsRow" + Convert.ToString(i % 2);
    }
    public static void clalert(string mess)
    {
        string s = "script";
        mess = mess.Replace("\\n", "§");
        mess = mess.Replace("\\", "\\\\");
        mess = mess.Replace("§", "\\n");
        K.Response.Write("<" + s + ">alert(\"" + mess + "\");</" + s + ">");
        Engine.theEngine.session.mo_alert = true;
    }
    public static decimal rptClauseID(string pclause)
    {
        object o1="",o2="",o3="",o4="",o5="",o6="",o7="",o8="",o9="",o10="",o11="",o12="",o13="",o14="",o15="",o16="";
        int l=pclause.Length,l2=0;
        Engine.theEngine.getmoconn();
        plWrapper.astrpt001 wrapp = new plWrapper.astrpt001();
        wrapp.Connection=Engine.theEngine.ADODBconnection;
        if(l<=2000)
            o1 =pclause.Substring(l2,l);
        else
        {
            o1 =pclause.Substring(l2,2000);
            l -=2000;
            l2+=2000;
            if(l<=2000)
                o2 =pclause.Substring(l2,l);
            else
            {
                o2 =pclause.Substring(l2,2000);
                l -=2000;
                l2+=2000;
                if(l<=2000)
                    o3 =pclause.Substring(l2,l);
                else
                {
                    o3 =pclause.Substring(l2,2000);
                    l -=2000;
                    l2+=2000;
                    if(l<=2000)
                        o4 =pclause.Substring(l2,l);
                    else
                    {
                        o4 =pclause.Substring(l2,2000);
                        l -=2000;
                        l2+=2000;
                        if(l<=2000)
                            o5 =pclause.Substring(l2,l);
                        else
                        {
                            o5 =pclause.Substring(l2,2000);
                            l -=2000;
                            l2+=2000;
                            if(l<=2000)
                                o6 =pclause.Substring(l2,l);
                            else
                            {
                                o6 =pclause.Substring(l2,2000);
                                l -=2000;
                                l2+=2000;
                                if(l<=2000)
                                    o7 =pclause.Substring(l2,l);
                                else
                                {
                                    o7 =pclause.Substring(l2,2000);
                                    l -=2000;
                                    l2+=2000;
                                    if(l<=2000)
                                        o8 =pclause.Substring(l2,l);
                                    else
                                    {
                                        o8 =pclause.Substring(l2,2000);
                                        l -=2000;
                                        l2+=2000;
                                        if(l<=2000)
                                            o9 =pclause.Substring(l2,l);
                                        else
                                        {
                                            o9 =pclause.Substring(l2,2000);
                                            l -=2000;
                                            l2+=2000;
                                            if(l<=2000)
                                                o10 =pclause.Substring(l2,l);
                                            else
                                            {
                                                o10 =pclause.Substring(l2,2000);
                                                l -=2000;
                                                l2+=2000;
                                                if(l<=2000)
                                                    o11 =pclause.Substring(l2,l);
                                                else
                                                {
                                                    o11 =pclause.Substring(l2,2000);
                                                    l -=2000;
                                                    l2+=2000;
                                                    if(l<=2000)
                                                        o12 =pclause.Substring(l2,l);
                                                    else
                                                    {
                                                        o12 =pclause.Substring(l2,2000);
                                                        l -=2000;
                                                        l2+=2000;
                                                        if(l<=2000)
                                                            o13 =pclause.Substring(l2,l);
                                                        else
                                                        {
                                                            o13 =pclause.Substring(l2,2000);
                                                            l -=2000;
                                                            l2+=2000;
                                                            if(l<=2000)
                                                                o14 =pclause.Substring(l2,l);
                                                            else
                                                            {
                                                                o14 =pclause.Substring(l2,2000);
                                                                l -=2000;
                                                                l2+=2000;
                                                                if(l<=2000)
                                                                    o15 =pclause.Substring(l2,l);
                                                                else
                                                                {
                                                                    o15 =pclause.Substring(l2,2000);
                                                                    l -=2000;
                                                                    l2+=2000;
                                                                    if(l<=2000)
                                                                        o16 =pclause.Substring(l2,l);
                                                                    else
                                                                    {
                                                                        o16 =pclause.Substring(l2,2000);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        return (Decimal)wrapp.get_Clauseid2(o1, o2, o3, o4, o5, o6, o7, o8, o9, o10, o11, o12, o13, o14, o15,o16);
    }
    public static void RStoJRS(ADODB.Recordset ors, string name, int formnum)
    {
	    RStoJRSX (ors,name,formnum,"false");
    }
    public static void RStoJRS2(ADODB.Recordset ors, string name, int formnum)
    {
	    RStoJRSX (ors,name,formnum,"true");
    }
    public static void RStoJRSX(ADODB.Recordset ors, string name, int formnum, string nosvc)
    {
	    char cr='\r'; //fld
        string sep;
        int n;

	    if (!ors.BOF || !ors.EOF)
		    ors.MoveFirst();
	    n=ors.Fields.Count;
	    Response.Write("var anames=new Array(");
	    sep="";
	    foreach(ADODB.Field fld in ors.Fields )
        {
		    Response.Write(sep + "\"" + fld.Name + "\"");
		    sep=",";
	    }
	    Response.Write(");" + cr);
	    Response.Write(name + "=new jRecordset(\"" + name + "\",anames," + formnum + "," + nosvc + ");" + cr);
    	
	    while(! ors.EOF)
        {
		    sep="";
		    Response.Write(name + ".addrow(new Array(");
	        foreach(ADODB.Field fld in ors.Fields )
            {
			    if (fld.Type==ADODB.DataTypeEnum.adDate)
                    try
                    {
				        Response.Write(sep + "\"" + K.getString(fld.Value) + "\"");
                    }
                    catch(Exception)
                    {
				        Response.Write(sep + "\"  \"");
                    }
			    else
				    Response.Write(sep + "\"" + EncodedString(K.getString(fld.Value)) + "\"");
                sep=",";
            }
		    Response.Write("));" + cr);
		    ors.MoveNext() ;
	    }
	    Response.Write(name + ".start();" + cr);
    }
    public static string EncodedString(string val)
	{
		if(val==null)return null;
		return Uri.EscapeDataString(val);
	}
    public static string DecodedString(string val)
	{
		if(val==null)return null;
		return Uri.UnescapeDataString(val);
	}

}
public struct Appdata
{
    public Transactions script;
    public string mo_site;
    public string message;
    public string mo_server;
    public DateTime mo_dtinfo;
}
public class Agent
{
    public static Agent theAgent = new Agent();

    public int mo_user=0;
    public string mo_mail;
    public string mo_emails;		
    public string mo_emailssel;		
    public string mo_mailsel;		
    public string mo_formatmail;	
    public string mo_formaturl;		
    public string mo_url;			
    public string mo_urlsel;		
    public string mo_feedback;		
    public string mo_feedbacksel;
    public string mo_feedbackechec;
    public string mo_feedbacksucces;
    public string mo_expediteur;	
    public string mo_impsel;		
    public string mo_attentesel;	
    public string mo_imprimante;	
    public int    mo_nbcopies;
    public string mo_objet;
    public string mo_poste;			

    private Hashtable droits;
    public Agent()
    {
        droits = new Hashtable();
    }
    public void addDroit(string d)
    {
        droits[d] = 1;
    }
    public void removeDroit(string d)
    {
        droits.Remove(d);
    }
    public bool hasDroit(string d)
    {
        return droits.ContainsKey(d);
    }
    public int getDroit(string d)
    {
        return droits.ContainsKey(d)?1:0;
    }
    public void setDroit(string d,int pdroit)
    {
        if(pdroit==0)
            removeDroit(d);
        else
            addDroit(d);
    }
}
public partial class Engine
{
    public static Engine theEngine;
    private Transaction transaction;
    private Transaction menutransaction;
    private Context ctx;
    public Stack callstack;
    private string wrname="PLWRAPPER";

    public Stack savestack;
    public bool noincr = false;

    public string SESSIONID;

    public string mo_npageout="0";

    public UnsavedSessionData session;
    public SavedSessionData ssession;

    public Appdata Application;

    public string site; //= "moniteurweb";
    public string langue = "undefined";//fr en

    public Hashtable wrappers;
    public int xxxtran;
    public Hashtable rform;
    public static bool bdomenu;

    public void init()
    {
        K.Session= HttpContext.Current.Session;
        K.Request= HttpContext.Current.Request;
        K.Response = HttpContext.Current.Response;
        K.Server= HttpContext.Current.Server;
    }

    public class UnsavedSessionData
    {
        public string mo_schema;
        public string mo_password;
        public string mo_instance;
        public string mo_application;
        public string ex_utilisateur;
        public string ex_utilisateur2;

        public Context mo_ctx;
        public string mo_return;
        //
        public string mo_result;
        public string mo_transaction;
        public string mo_caller;
        public Exception mo_errobj;
        public int mo_viaxml;
        public bool mo_clear;
        public bool mo_messext;
        public int mo_save_callstack;
        public int mo_save_savestack;
        public bool mo_messlu;
        public string mo_modal;
        public string mo_menu;
        public bool mo_intrans;
        public string mo_errmess;
        public string mo_nompage;
        public string mo_request;
        public string mo_langue;
        public bool mo_haserror;
        public bool mo_noalerterror;
        public string mo_server;
        public string mo_appli;
        public string mo_startmenu;
        public string mo_mode;
        public string mo_flavor;
        public bool mo_alert;
        public string mo_passwd;
        public int mo_user_chef;
        public int mo_telmonitor;
        public string mo_lastnum;
        public int init;

        public MSXML2.XSLTemplate mo_xsl, mo_xsl2;
    }
    public struct SavedSessionData
    {
        public string mo_title;
        public string mo_nextcomment;
        public string mo_commandes;
        public bool mo_nosave;
        public string mo_lastreturn;
        public string mo_old_menu;
        public MSXML2.IXMLDOMDocument2 mo_xmls, mo_xmls2;
        public bool breinit;
        public Agent mo_agent;

        public void save(Stack s)
        {
            mo_agent = Agent.theAgent;
            s.Push(this);
        }
        public void restore(Stack s)
        {
            this = (SavedSessionData)s.Pop();
            Agent.theAgent = mo_agent;
        }
    }


    public class Connection
    {
        public  const bool mo_persist_connection=true;

        public string constr;
        public int nbusers;// nb utilisateurs
        public bool isintran;// indic transac en cours
        public int translevel;// niveau transaction
        public ArrayList transconn; // participants à la transaction
        public ADODB.Connection cnx;
        public OracleTransaction trans;

        public Connection(string pconstr)
        {
            constr = pconstr;
            nbusers = 0;
            isintran = false;
            translevel = 0;
            transconn = new ArrayList();
        }
        public ADODB.Connection getconnection()
        {
            if (nbusers>0)
            {
                if(!mo_persist_connection)
                {
                    nbusers++;
                    return cnx;
                }
            }
            else
            {
                nbusers=1;
                cnx=new ADODB.Connection();
                cnx.ConnectionString=constr;
                cnx.Open(constr,null,null,0);                
            }
            return cnx;
        }
        public void closeconnection()
        {
            if (!mo_persist_connection && nbusers > 0)
            {
                try
                {
                    trans.Commit();
                }
                catch (Exception) { }
                cnx.Close();
                nbusers = 0;
            }
        }
    }
    private Hashtable connections;
    public void addconnection(string pname, string pconstr)
    {
        connections.Add(pname, new Connection(pconstr));
    }
    public ADODB.Connection getconnection(string pname)  //ADODB.Connection
    {
        ADODB.Connection cnx=((Connection)connections[pname]).getconnection();
        gloinit(cnx);
        //connectionobj = cnx;
        ADODBconnection = cnx;
        return cnx;
    }
    //public object connectionobj;
    public ADODB.Connection ADODBconnection;
    public void getmoconn()
    {
        if (ADODBconnection == null)
            getconnection("MO_CONN");
    }
    public void closeconnections()
    {
        foreach (Connection cnx in connections.Values)
            cnx.closeconnection();
    }
    public bool isintrans()
    {
        foreach (Connection cnx in connections.Values)
            if (cnx.isintran)
                return true;
        return false;
    }
    public void xmltransform()
    {
        xmlService.ParameterList plang = new xmlService.ParameterList();
        string s1="lang";
        object o2 = session.mo_langue;
        plang.Add(ref s1, ref o2);
        object o1 = session.mo_xsl;
        o2 = plang;
        K.Response.Write(XmlServiceFactory2().service.layout.ApplyStylesheet(ref o1, ref o2));
        o1=null;
        o2=null;
        plang=null;
    }
    public void error(string mess)
    {
        K.Response.Write("erreur interceptée par le moniteur<br>" + mess + "<br>");
        K.Session["MO_ERROR"]=mess;
        K.Server.Transfer(getpath("asp/MO_ERROR.ASP"));
    }
    public xmlService.Document2 XmlServiceFactory2()
    {
        xmlService.Document2 xmls = new xmlService.Document2();	
	    xmls.XmlDocument=(MSXML2.FreeThreadedDOMDocument) ssession.mo_xmls;	
	    return xmls;
    }
    
    private void doabandonappli()
    {
        this.mdl_initcss();
        K.Response.Clear();
        K.Session.Abandon();
        K.Response.End();
    }

    /// <summary>
    /// classe Context ...
    /// </summary>
    public class Context
    {
        public Engine engine;
        public Transaction transaction;
        public int ptr;
        public string caller;
        public bool brequest;
        public bool btransfer;
        public bool bissaved;
        public Stack stack;
        public Stack laststack;
        public int ptrlastpage;
        public int ptrlastaff;

        public Appdata Application;

        public Engine.UnsavedSessionData session;
       // public Engine.SavedSessionData ssession;
        public Stack savestack;
        public Stack callstack;

        public Context() { }

        public Context(Engine e,Transaction t, string c)
        {
            engine = e;
            transaction = t;
            ptr = 0;
            stack = new Stack();
            laststack = new Stack();
            ptrlastpage = -1;
            ptrlastaff = -1;
            brequest = false;
            btransfer = false;
            bissaved = false;
            caller = c;

            Application=engine.Application;

            session=engine.session;
            //ssession=engine.ssession;
            savestack=engine.savestack;
            callstack=engine.callstack;
        }
        /// <summary>gestion erreurs</summary>
        /// <param name="mess"> message</param>
        public void error(string mess)
        {
            engine.error(mess);
        }

        public Instruction getInstruction()
        {
            if (ptr >= transaction.instructions.Count)
                return null;
            return (Instruction)transaction.instructions[ptr++];
        }
        private void skip()
        {
            int i,idep,nbpar;
            enInstrType t;
            
            i = ptr;
            idep = i;
            nbpar = 0;

            while (true)
            {
                t=((Instruction)(transaction.instructions[i++])).type;
                switch (t)
                {
                    case enInstrType.BLOCK:
                        nbpar++;
                        break;
                    case enInstrType.ENDBLOCK:
                        nbpar--;
                        break;
                    case enInstrType.REPEAT:
                        continue;
                    case enInstrType.CASE:
                        continue;
                    case enInstrType.OTHER:
                        continue;
                    case enInstrType.CALL:
                        continue;
                }
                if (nbpar == 0)
                    break;
            }
            ptr = i;
        }
        private void save()
        {
            engine.ssession.save(savestack);
        }
        public void restore()
        {
            engine.ssession.restore(savestack);
        }
        public void play()
        {
            Transaction ntransaction;
            Instruction instruction;
            Connection cn,cn2;
            Ctxinstr ct;
            bool mo_skipcase = false,mo_more=true;    // saut CASE
            Context tmpctx = engine.ctx;
            string mo_rc="",t,p,c,lang,langa,src,str,commandes,page,tmp,txt,x="",name;
            bool not,bfound,btrue,r,bcon2=false,isin=false;
            int i, mpos;
            object obj,obj2,wrapper,ors;
            object[] arg=new object[1];
            xmlService.Document2 xmls;
          
            ArrayList arr,ltr;
            Regex rx;

            Transaction trans;
            Context nctx;
            ADODB.Connection cnx;
            OracleCommand cmd;
            plWrapper.astext001 wrp;

            xmlService.Document2 objXMLS;
            MSXML2.IXMLDOMDocument2 tmpxmls;
            MSXML2.XSLTemplate oxsl;
            MSXML2.FreeThreadedDOMDocument tmpxmldom;
            xmlService.ParameterList plang;
            
            engine.ctx = this;
            session.mo_transaction = transaction.name;
            session.mo_caller = caller;

            while (mo_more)
            {
                tmp = session.mo_return.ToUpper();
                if(tmp!="" && tmp!=null)
                {	// si le code retour est à blanc, mo_rc n'est pas modifié
                      if (tmp!="ERROR")
		                    engine.ssession.mo_lastreturn=tmp;    
                      i=tmp.Length;
                      while(--i>0)		// élimination blancs et ;
                        if((c=tmp.Substring(i,1))!=" " && c!=";")
                            break;
                      tmp=tmp.Substring(0,i+1);
                      mo_rc=tmp;

                      switch (mo_rc) 
                      {
                          case "ERROR":		// erreur : retour vers la dernière page affichée
                              if(ptrlastaff<0)
                                    return;		// retour script supérieur
                              ptr=ptrlastaff;
                                   //mo_ctx("stack") est la pile d'exécution
                                   //mo_ctx("laststack") est la sauvegarde correspondante à la dernière page affichée
                              stack=(Stack)laststack.Clone();
                              break;

                          case "LOOP":	// retour en début de boucle REPEAT
                          case "BREAK":     // fin de boucle
                                // on recherche dans la pile d'exécution un ordre REPEAT
                                bfound=false;
                                while(stack.Count>0)
                                {
                                    ct=(Engine.Ctxinstr)stack.Pop();
                                    if((enInstrType)ct.verb==enInstrType.REPEAT)
                                    {
                                        ptr=(int)ct.data;  // on se positionne en début de boucle
                                        if (mo_rc == "BREAK")   // si BREAK
                                        {
                                            skip();    // on saute l'ordre REPEAT
                                            fininstr(ref mo_skipcase);
                                        }
                                        bfound=true;
                                        break;
                                    }  
                                }
                                if(!bfound)
                                    return;			// non trouvé on sort
                                break;
                          case "ERRSYS" :
                                error("6 erreur systeme");
                                break;
                          case "ABORT" :  // arrêt transaction en cours
                          case "RETURN" :
                                mo_more=false;
                                if(transaction.name!="request")
                                {
                                  session.mo_return="";  
                                  mo_rc="";
                                }
                                mo_skipcase = false;
                                continue;
                          case "ABANDON" :  // arrêt session
			                    engine.doabandonappli();
                                goto case "ENDMODAL";
		                  case "ENDMODAL" :  // fin fenêtre modale
		                        if(session.mo_modal!="NO")
		                        {
				                    if(callstack.Count>0)
				                    {
					                    session.mo_ctx=(Context)callstack.Pop(); 
					                    if(session.mo_ctx.bissaved)
					                    {
						                    restore();
					                        session.mo_ctx.bissaved=false;
					                    }
					                    session.mo_transaction =session.mo_ctx.transaction.name;
					                    session.mo_caller =session.mo_ctx.caller;
				                    }		    
				                    session.mo_modal="NO";
                                    K.Response.Write("<script>document.cookie = 'modal=NO;path='+escape('/');document.cookie = 'result=" + Uri.EscapeDataString((string)K.Session["MO_RESULT"]) + ";path='+escape('/');window.close();</script>");
				                    engine.mo_npageout="-1";
				                    K.Response.End();
			                    }
			                    break;

                          default :
			                    c=mo_rc.Substring(0,1);
			                    if(c==">")			//	><nom transaction>   (TRANSFER)
			                    {
				                    t=mo_rc.Substring(1);
				                    session.mo_menu=t;												
                                    ntransaction = engine.transactions.get(t);
					                //error("7 Transaction "+t+" inexistante");
                                    nctx = new Context(engine,ntransaction, "");
				                    nctx.btransfer=true;
				                    nctx.brequest=brequest;
				                    session.mo_ctx=nctx;  // le nouveau contexte est conservé dans la session				
				                    session.mo_return=""; 

				                    nctx.play();  // on l'exécute en récursif .   Au cas où l'on revient : 
				                    
                                    if(nctx.brequest)
					                    callstack.Pop();
				                    if(callstack.Count>0)
					                    callstack.Pop(); // on dépile l'ancien contexte
				                    if(bissaved)
				                    {
					                    restore();
					                    bissaved=false;
				                    }       
				                    session.mo_transaction =transaction.name;
				                    session.mo_caller =caller;
				                    session.mo_ctx=this;    // on le stocke dans la session
			                    }		    
			                    else if(c=="@")			//	@<nom transaction>   (BRANCH)
			                    {
				                    t=mo_rc.Substring(1);
				                    session.mo_menu=t;												
                                    ntransaction = engine.transactions.get(t);
					                //error("7 Transaction "+t+" inexistante");
                                    nctx = new Context(engine,ntransaction, "");
				                    session.mo_ctx=nctx;  // le nouveau contexte est conservé dans la session				
				                    session.mo_return=""; 
                                    engine.savestack=new Stack();
                                    engine.callstack=new Stack();
                                    savestack=engine.savestack;
                                    callstack=engine.callstack;

				                    nctx.play();  // on l'exécute en récursif .   Au cas où l'on revient : 

                                    K.Session.Abandon();
				                    K.Response.Clear();
				                    K.Response.End();			
			                    }		    
			                    else if(c=="=")			//	=<nom transaction>   (CALL)
			                    {
				                    t=mo_rc.Substring(1);
				                    engine.ssession.mo_old_menu=session.mo_menu;												
				                    session.mo_menu=t;												

                    // les instructions marquées CALL servent a appeler t par "CALL TRANS t" au lieu de "BRANCH"

				                    bissaved=true;  // CALL
                                    stack.Push(new Ctxinstr(enInstrType.CALL,ptrlastpage));
		                            engine.ssession.mo_lastreturn="";    // CALL
		                            engine.ssession.breinit=true;				// testé au retour pour pgl qui lit directement Request.Form
				                    save();		// CALL

                                    ntransaction = engine.transactions.get(t);
					                //error("7 Transaction "+t+" inexistante");
                                    nctx = new Context(engine,ntransaction,session.mo_transaction );
				                    session.mo_ctx=nctx;  // le nouveau contexte est conservé dans la session				
                                    callstack.Push(this);
				                    session.mo_return=""; 

				                    nctx.play();  // on l'exécute en récursif .   Au cas où l'on revient : 
				                    
				                    if(callstack.Count>0)
					                    callstack.Pop(); // on dépile l'ancien contexte
				                    if(bissaved)
				                    {
					                    restore();
					                    bissaved=false;
				                    }       
				                    session.mo_transaction =transaction.name;
				                    session.mo_caller =caller;
				                    session.mo_ctx=this;    // on le stocke dans la session
			                    }
                                break;
                      }
                }
                session.mo_return="";  // raz code retour. mo_rc est conservé

                mpos = ptr;
                instruction = getInstruction();
                if (instruction == null)
                    break;

                switch (instruction.type)
                {
                    case enInstrType.LANGUAGE:  // langue
                        session.mo_langue = instruction.getparam(0);		// code langue
                        break;

                    case enInstrType.ENDMODAL:  // fin fenêtre modale
                        if (session.mo_modal != "NO")
                        {
                            if (callstack.Count > 0)
                            {
                                session.mo_ctx = (Context)callstack.Pop();
                                if (session.mo_ctx.bissaved)
                                {
                                    restore();
                                    session.mo_ctx.bissaved = false;
                                }
                                session.mo_transaction = session.mo_ctx.transaction.name;
                                session.mo_caller = session.mo_ctx.caller;
                            }
                            session.mo_modal = "NO";
                            K.Response.Write("<script>document.cookie = 'modal=NO;path='+escape('/');document.cookie = 'result=" + Uri.EscapeDataString((string)K.Session["MO_RESULT"]) + ";path='+escape('/');window.close();</script>");
                            engine.mo_npageout = "-1";
                            K.Response.End();
                        }
                        break;

                    case enInstrType.ABANDON:
                        engine.doabandonappli();
                        break;

                    case enInstrType.ABORT:
                    case enInstrType.RETURN:
                        mo_more = false;
                        if (transaction.name != "request")
                            session.mo_return = "";
                        break;

                    case enInstrType.REPEAT:
                        stack.Push(new Ctxinstr(enInstrType.REPEAT, mpos));
                        mo_skipcase = false;
                        continue;

                    case enInstrType.LOOP:
                    case enInstrType.BREAK:
                        bfound = false;
                        while (stack.Count > 0)
                        {
                            ct = (Ctxinstr)stack.Pop();
                            if (ct.verb == enInstrType.REPEAT)
                            {
                                bfound = true;
                                ptr =(int) ct.data;
                                if (instruction.type == enInstrType.BREAK)
                                {
                                    skip();   
                                    fininstr(ref mo_skipcase);
                                }
                                break;
                            }
                        }
                        if (!bfound)
                        {
                            session.mo_return = Language.getVerb(instruction.type);
                            return;
                        }
                        break;

                    case enInstrType.CASE:
                        if (mo_skipcase)
                        { // si l'on est déja tombé dans un CASE, on saute jusqu'à un
                            skip();   // ordre différent de CASE ou OTHER
                            fininstr(ref mo_skipcase);
                            continue;
                        }

                        not = instruction.haskeyword(0); // NOT
                        c = instruction.getparam(0);
                        if (c == "EMPTY")
                            c = "";
                       
                        if (c.StartsWith("DR:"))
                        {
                            if (!Agent.theAgent.hasDroit(c.Substring(3)) ^ not)
                            { // le dictionnaire contient 1 si droit = oui
                                skip();			//  non autorisé
                                fininstr(ref mo_skipcase);
                                continue;
                            }
                        }
                        else					// code retour
                            if ((c != mo_rc) ^ not)
                            {
                                skip();		// test négatif
                                fininstr(ref mo_skipcase);
                                continue;
                            }
                        stack.Push(new Ctxinstr(enInstrType.CASE, null));
                        continue;

                    case enInstrType.OTHER:
                        if (mo_skipcase)
                        { // si l'on est déja tombé dans un CASE, on saute jusqu'à un
                            skip();   // ordre différent de CASE ou OTHER
                            fininstr(ref mo_skipcase);
                            continue;
                        }
                        stack.Push(new Ctxinstr(enInstrType.OTHER, null));
                        continue;

                    case enInstrType.BLOCK:
                        mo_skipcase = false;
                        stack.Push(new Ctxinstr(enInstrType.BLOCK, null));
                        continue;

                    case enInstrType.ENDBLOCK:
                        stack.Pop();
                        break;

                    case enInstrType.CALL:
                        mo_skipcase = false;
                        bissaved = true;
                        stack.Push(new Ctxinstr(enInstrType.CALL, null));
                        save();
                        continue;

                    //---------------------------------DISP page----------------------------------------- DISP
                    case enInstrType.DISP:
                        ptrlastpage = ptr;
                        ptrlastaff = ptr;
                        laststack =(Stack) stack.Clone();
                        break;

                    case enInstrType.BRANCH:
                    case enInstrType.TRANSFER:
                    case enInstrType.TRANS:
                        t = instruction.getparam(0);   // lecture du nom
                        trans = engine.transactions.get(t);

                        if (instruction.type == enInstrType.TRANS)
                            nctx = new Context(engine, trans, transaction.name);
                        else
                            nctx = new Context(engine, trans, "");

                        if (instruction.type == enInstrType.TRANS)
                            callstack.Push(this);
                        else 
                            if (instruction.type == enInstrType.TRANSFER)
                            {
                                nctx.btransfer = true;
                                nctx.brequest  = brequest;
                            }
                        session.mo_ctx = nctx;
                        if (instruction.type == enInstrType.BRANCH)
                        {
                            savestack = new Stack();
                            callstack = new Stack();
                        }

                        nctx.play();  // on l'exécute en récursif .   Au cas où l'on revient : 
                        
                        if (instruction.type == enInstrType.BRANCH)
                        {
                            K.Session.Abandon();
                            K.Response.Clear();
                            K.Response.End();
                        }
                        if (nctx.brequest && nctx.btransfer)
                            callstack.Pop();
                        if (callstack.Count > 0)
                            callstack.Pop();
                        if (bissaved)
                        {
                            restore();
                            bissaved = false;
                        }
                        session.mo_transaction = transaction.name;
                        session.mo_caller = caller;
                        mo_skipcase = false;
                        break;

                    case enInstrType.EXECP:
                    case enInstrType.EXEC:
                        bdomenu = true;
                        p = engine.getpath(instruction.getparam(0));
                        ptrlastpage = mpos;
                        if (p.StartsWith("mdl_"))
                        {
                            Type tpe = typeof(Engine);
                            tpe.InvokeMember(p, BindingFlags.InvokeMethod, null, engine, null);
                        }
                        else
                        {
                            session.mo_nompage = p;
                            K.Server.Execute(p);
                        }
                        break;
                    //---------------------------ALERT "message"-----------------------------------
                    case enInstrType.ALERT:
                        c = instruction.getparam(0);
                        K.Response.Write("<script>alert(unescape('"+Uri.EscapeDataString(c)+"'));</script>");
                        break;

                    //---------------------------TITLE "titre"-----------------------------------
                    case enInstrType.TITLE:
                        c = instruction.getparam(0);
                        lang = instruction.getparam(1);
                        langa = engine.langue.ToUpper();
                        if (lang == null && langa == "UNDEFINED" || langa == lang)
                        {
                            K.Response.Write("<html><head><title>NASTER V1.0 " + c + "</title></head>");
                            engine.ssession.mo_title = c;
                        }
                        break;

                    //---------------------------WRNAME <nom wrapper>-----------------------------------
                    case enInstrType.WRNAME:
                        engine.wrname = instruction.getparam(0);
                        break;

                    //---------------------------DEBUG-------------------------------------------
                    case enInstrType.BP:
                        break;

                   //---------------------------SQLTRACE <connection> TRUE/FALSE   trace oracle -----------------------------------------
	                case enInstrType.SQLTRACE:
                        /*
                        src=instruction.getparam(0);		// nom datasource
                        cnx=engine.getconnection(src);
		                btrue=instruction.haskeyword(0);			// TRUE  /  FALSE
		                cmd=new OracleCommand("alter session set sql_trace = "+(btrue?"true":"false"),cnx);
                        cmd.ExecuteNonQuery();
		                cnx.Close();
                        //cnx.Dispose();
		                cmd.Dispose();
                         */
		                break;

                    //---------------------------------COMMANDES "libellé=option/..." par langue ---------------COMMANDES
                    case enInstrType.COMMANDES:
                        commandes = instruction.getparam(0);
                        lang = instruction.getparam(1);
                        if (lang == null || engine.langue.ToUpper() == lang)
                            engine.ssession.mo_commandes = commandes;
                        break;
                    //---------------------------------NOCOMMANDES  --------------------------------------------NOCOMMANDES
                    case enInstrType.NOCOMMANDES:
                        engine.ssession.mo_commandes =  null;
                        break;
                    
                    //---------------------------------LOADXML <fic>.xml-----------------------------------LOADXML
	                case enInstrType.LOADXML:
                        objXMLS = new xmlService.Document2();
                        page = engine.getpath(instruction.getparam(0));
                        page = K.Server.MapPath(page);	    
	                    try
                        {
                            obj = page;
	                        objXMLS.LoadXML(ref obj);
                            objXMLS.service.Action="moniteur.aspx";
                            engine.ssession.mo_xmls=objXMLS.XmlDocument;
		                }
                        catch(Exception e)
                        {
                            error("10 échec(c) load xml "+page + "\n\r" + e);
                        }
		                break;
                    //---------------------------------SWAPXML -----------------------------------SWAPXML
	                case enInstrType.SWAPXML:
	                    tmpxmls=engine.ssession.mo_xmls;
	                    engine.ssession.mo_xmls=engine.ssession.mo_xmls2;
	                    engine.ssession.mo_xmls2=tmpxmls;
	                    tmpxmls=null;
	                    break;                        
                    //---------------------------------CLOSEXML ----------------------------------------- CLOSEXML
   	                case enInstrType.CLOSEXML:
		                break;		
                    //---------------------------------TRANSFORM [<fic>.xsl,langue]----------------------------TRANSFORM
	                case enInstrType.TRANSFORM:
		                page=instruction.getparam(0);
                        lang = instruction.getparam(1);
                	    
	                    if(page!=null)
	                    {
                              page = K.Server.MapPath(page);
	                          tmpxmldom=new MSXML2.FreeThreadedDOMDocument();
                              r=tmpxmldom.load(page);
                              if(!r)
	                            error("11 échec load xml "+page);
                              oxsl=session.mo_xsl2;
		                      oxsl.stylesheet = tmpxmldom;
		                      tmpxmldom=null;
                              lang=instruction.getparam(2).ToLower();
                        }
                        else
                              oxsl=session.mo_xsl;
                        
                        plang=new xmlService.ParameterList();
                        obj=lang;
                        c="lang";
                        plang.Add(ref c,ref obj);
                        try
                        {
                            obj = oxsl;
                            obj2 = plang;
                            K.Response.Write(engine.XmlServiceFactory2().service.layout.ApplyStylesheet(ref obj,ref obj2));
                            obj = null;
                            obj2 = null;
	                    }
                        catch(Exception e)
                        {
                            session.mo_errobj=e;
                            error("12 TRANSFORM");
                        }
                        session.mo_viaxml=1;
		                break;

        //----------------------------------SEND [nosave][clear][noreturn][nocontent]--------------------------------------- SEND        
                    case enInstrType.SEND :		 // envoi de la page
                        engine.wrappers.Clear();
		                engine.ssession.mo_nosave=instruction.haskeyword(0); // pas de sauvegarde forme
		                session.mo_clear=instruction.haskeyword(1); // clear xmls

                        if (Agent.theAgent.mo_user!=0 && session.mo_messext && Agent.theAgent.hasDroit("BEXT"))
		                {
                            wrp = new plWrapper.astext001();
                            engine.getmoconn();
                            wrp.Connection= engine.ADODBconnection;
                            obj = Agent.theAgent.mo_user;
			                if((int)wrp.hasMsggestionnaire(ref obj)==1)
			                {
				                K.Response.Write("<script>alert('il y a des messages EXTRANET');</script>");
				                session.mo_messext=false;
			                }
			                wrp=null;		
			                cnx=null;
		                }
		                engine.closeconnections();
                        
                        K.Response.Write("<script>var sMO_TRANSACTION='"+session.mo_transaction+"';var sMO_NOMPAGE='"+session.mo_nompage+"';</script>");
                        
                        if(instruction.haskeyword(2)) // noreturn
		                {
			                if(session.mo_save_callstack<callstack.Count)
			                {
                                Utility.setcount(callstack, session.mo_save_callstack + 1);
			                    session.mo_ctx=(Context)callstack.Pop();
				                session.mo_transaction =session.mo_ctx.transaction.name;
				                session.mo_caller =session.mo_ctx.caller;
			                }
			                if(session.mo_save_savestack<savestack.Count)
			                {
				                Utility.setcount(savestack,session.mo_save_savestack+1);
				                restore();
			                }
		                }
                        if(instruction.haskeyword(3)) // nocontent
		                {
		                    K.Response.Status="204 NO CONTENTS";	//retour sans contenu : n'efface pas la page
		                    K.Response.End();
		                }

		                if (Application.message !=null)
		                {
			                if(!session.mo_messlu)
			                {
				                K.Response.Write("<script>alert('"+Application.message+"');</script>");
				                session.mo_messlu=true;
			                }
		                }
		                else
			                session.mo_messlu=false;
                		
		                if(session.mo_viaxml==1 && !engine.ssession.mo_nosave)
			                session.mo_viaxml=2;
		                else
			                session.mo_viaxml=0;

                        K.Response.End();
                         
                        break;

                    //------------------------------SAVE-------------------------------------------- SAVE
                    case enInstrType.SAVE:     // sauvegarde de l'objet Session dans SAVESTACK
                        save();
                        break;

                    //----------------------------RESTORE---------------------------------------------- RESTORE
                    case enInstrType.RESTORE:  // restauration Session
                        restore();
                        break;

                    //-------------------------------EVAL script sans mise résultat dans MO_RETURN---- §  §
                    case enInstrType.EVAL2:     // évaluation javascript  §
                        txt = instruction.getparam(0);    // lecture du script
                        //try
                       // {
                            Evaluator.EvalNoResult(txt);
                        //}
                        //catch (Exception)
                        //{
                        //    error("14 erreur dans l'exécution du script: \n" + txt);
                        //}
                        mo_skipcase = false;
                        break;

                    //-------------------------------EVAL script avec mise résultat dans MO_RETURN----- EVAL
                    case enInstrType.EVAL:     
                        txt = instruction.getparam(0);    // lecture du script
                        try
                        {
                            x=Evaluator.EvalToString(txt);
                        }
                        catch (Exception)
                        {
                            error("15 erreur dans l'exécution du script: \n" + txt);
                        }
                        if (x != null && x!="")
                        {
                            try { session.mo_return = x; } // résultat dans code retour
                            catch (Exception) { session.mo_return = ""; }
                        }
                        break;

            //---------------------------------------------------------------------------------------------
            //---------------------------------------------- CONNECTIONT ET TRANSACTIONS
            //---------------------------------------------------------------------------------------------
                    
            //---------------------------------DEFINEDATASOURCE nom="chaine"-------------------------DEFINEDATASOURCE
                    case enInstrType.DEFINEDATASOURCE :
                        src=instruction.getparam(0);		// nom datasource
                        str=instruction.getparam(1);		// chaine de connection
                        txt=session.mo_schema;
                        if(src=="MO_CONN" && txt!="")
                        {
                            rx= new Regex(@"User Id=\w*;", RegexOptions.IgnoreCase);
                            str=rx.Replace(str,"User Id="+txt+";");

                            rx= new Regex(@"Password=\w*;", RegexOptions.IgnoreCase);
			                txt=session.mo_password;
                            str = rx.Replace(str, "Password=" + txt + ";");

                            rx= new Regex(@"Data Source=\w*;", RegexOptions.IgnoreCase);
			                txt=session.mo_instance;
                            str = rx.Replace(str, "Data Source=" + txt + ";");
		                }                        
                        engine.addconnection(src,str);
                        break;

//---------------------------------GETCONNECTION nom [,wrapper,..]----------------------------------------GETCONNECTION 
                    case enInstrType.GETCONNECTION2 :		// pour 2eme connexion = CR@encours= impression
                        bcon2=true;
                        goto case enInstrType.GETCONNECTION;
                    case enInstrType.GETCONNECTION :
                        src=instruction.getparam(0);		// nom datasource
                        // recherche d'une connection disponible
                        cn=(Connection)engine.connections[src];
                        if(cn==null)
                          error("16 erreur connexion inexistante");
                        
                        cnx=cn.getconnection();
                        arg[0]=cnx;
                        
                        arr=instruction.getlist(1);  // liste des arguments (wrappers)
                        foreach (string lstr in arr)
                        {
			                name="WR_"+lstr;
                            obj=engine.wrappers[name];
                            if (obj == null)
                            {
                                obj = K.Server.CreateObject(engine.wrname + "." + lstr);
                                engine.wrappers[name] = obj;
                            }
                            obj.GetType().InvokeMember("Connection", BindingFlags.SetProperty, null, obj, arg);
                        }  
                        cnx=null;
                        break;

                    //---------------------------------GLOINIT nom ----------------------------------------GETCONNECTION 
                    case enInstrType.GLOINIT:
                        src = instruction.getparam(0);		// nom datasource
                        cn=(Connection)engine.connections[src];
                        cnx = cn.getconnection();

                        engine.SESSIONID = engine.gloinit(cnx);	// initialisation du package GLO
                        cnx = null;
                        break;

                    //---------------------------------CLOSECONNECTION nom-----------------------------------------CLOSECONNECTION 
                    case enInstrType.CLOSECONNECTION2:		// pour 2eme connexion non persistante
                        bcon2 = true;
                        goto case enInstrType.CLOSECONNECTION;
                    case enInstrType.CLOSECONNECTION:
                        src = instruction.getparam(0);		// nom datasource
                        if (!Connection.mo_persist_connection || bcon2)
                        {
                            cn=(Connection)engine.connections[src];
                            if (cn.nbusers > 0)
                                break;
                            cnx = cn.getconnection();
                            cnx.Close();
                            cnx = null;
                            cn.isintran = false;
                            session.mo_intrans = engine.isintrans();
                        }
                        mo_skipcase = false;
                        break;
            //-----------------------------BEGINTRANS [datasource,..]---------------------------------------BEGINTRANS
                    case enInstrType.BEGINTRANS :    // transaction SGBD
                        arr=instruction.getlist(0);  
                        foreach (string lsrc in arr)
                        {
                            cn = (Connection)engine.connections[lsrc];
                            if(cn.nbusers>0 && cn.translevel++==0)
                            {
                                cnx = cn.getconnection();
                                //cn.trans = cnx.BeginTransaction();                              
                                cnx.BeginTrans();              
                                engine.xxxtran++;            
                                cn.isintran=true;	// indic transac en cours
                                cn.transconn=(ArrayList)arr.Clone() ;  // connexions participant a la transaction
                          }
                        }  
                        session.mo_intrans=true;
                        break;

            //----------modifié-------------COMMIT [datasource,..]-------------------------------------------- COMMIT
                    case enInstrType.COMMIT :
                        arr=instruction.getlist(0);  
                        foreach (string lsrc in arr)
                        {
                            cn = (Connection)engine.connections[lsrc];
                            if (cn.nbusers > 0 && (cn.translevel == 0 || --cn.translevel == 0))  
	                        {
	                            isin=true;
	                            ltr=cn.transconn;
	                            foreach(string lstr in ltr)
	                                if(!arr.Contains(lstr))
	                                {
	                                    isin=false;
	                                    break;
	                                }
                	            
                	            
	                            if(cn.isintran && isin) // si pas d'autre connection
	                            {
				                    try
                                    {
                                        cnx = cn.getconnection();
                                        //cn.trans.Commit();
                                        cnx.CommitTrans();
                                        engine.xxxtran--; 
                                    } 
				                    catch(Exception){}
	                                cn.isintran=false;	// indic transac en cours
	                            }
	                        }
	                    }  
	                    session.mo_intrans=engine.isintrans();
	                    break;

            //-------------------------------ROLLBACK [datasource,..]------------------------------------------ ROOLBACK
                    case enInstrType.ROLLBACK :
                        arr=instruction.getlist(0);  
                        foreach (string lsrc in arr)
                        {
                            cn = (Connection)engine.connections[lsrc];
                            if (cn.nbusers > 0)
	                        {
                                if(cn.isintran)
                                {
			                        try
                                    {
                                        cn.trans.Rollback(); 
                                        engine.xxxtran--;
                                    } 
                                    catch(Exception){}
                                    cn.isintran=false;	// indic transac en cours
                                    // raz niveau imbrication
                                    cn.translevel=0;
			                        logmess("case ROLLBACK"," ROLLBACK");
                                }
                            }

	                        ltr=cn.transconn;
                            foreach (string key2 in ltr)
                            {
                                cn2 = (Connection)engine.connections[key2];
                                if (cn2.nbusers > 0)
                                {
                                    if (cn2.isintran)
                                    {
                                        try
                                        {
                                            cn2.trans.Rollback();
                                            engine.xxxtran--;
                                        }
                                        catch (Exception) { }
                                        cn2.isintran = false;	// indic transac en cours
                                        // raz niveau imbrication
                                        cn2.translevel = 0;
                                    }
                                }

                            }
	                    }  
	                    session.mo_intrans=engine.isintrans();
	                    break;

//---------------------------------------------------------------------------------------------
//----------------------------- WRAPPER
//---------------------------------------------------------------------------------------------


                    //---------------------------- GETWRAPPER wrapper1,..---------------------------------------GETWRAPPER
                    case enInstrType.GETWRAPPER:    // création plwrapper.<wrapper1>
                        arr = instruction.getlist(0);
                        foreach (string lstr in arr)
                        {
                            name = "WR_" + lstr;
                            if (engine.wrappers[name] == null)
                                engine.wrappers[name] = K.Server.CreateObject(engine.wrname + "." + lstr);
                        }
                        break;
                      
                    //---------------------------- CLOSEWRAPPER wrapper1,..---------------------------------------GETWRAPPER
                    case enInstrType.CLOSEWRAPPER:    // création plwrapper.<wrapper1>
                        arr=instruction.getlist(0);  
                        foreach(string lstr in arr)
                        {
			                name="WR_"+lstr;
                            engine.wrappers[name] = null;
                        }  
                        break;
                    //---------------------------- UPDATE ---------------------------------------
                    case enInstrType.UPDATE :    // appel de la méthode Modifie d'un wrapper
                        wrapper=engine.wrappers["WR_"+instruction.getparam(0)];		// nom wrapper
                        ors = wrapper.GetType().InvokeMember("RecordStructure", BindingFlags.GetProperty, null, wrapper, null); 
                        try
                        {
		                    xmls=engine.XmlServiceFactory2();
                            obj=xmls.GetType().InvokeMember("Service",BindingFlags.GetProperty, null, wrapper, null);
                            arg[0]=ors;
                            obj.GetType().InvokeMember("ElementsToRecordset",BindingFlags.InvokeMethod,null,obj,arg);
		                }
                        catch(Exception e)
                        {
                            session.mo_errobj=e;
                            engine.error("17 update");
                        }		
		                try
                        {
                            arg[0]=ors;
                            wrapper.GetType().InvokeMember("Modifie", BindingFlags.InvokeMethod, null, wrapper, arg);
		                }
                        catch(Exception){}
                        break;

                    //---------------------------- DELETE ---------------------------------------
                    case enInstrType.DELETE:    // appel de la méthode Delete d'un wrapper (METHODE NON GENEREE automatiquement)
                        wrapper = engine.wrappers["WR_" + instruction.getparam(0)];		// nom wrapper
                        ors = wrapper.GetType().InvokeMember("RecordStructure", BindingFlags.GetProperty, null, wrapper, null);
                        try
                        {
                            xmls = engine.XmlServiceFactory2();
                            obj = xmls.GetType().InvokeMember("Service", BindingFlags.GetProperty, null, wrapper, null);
                            arg[0] = ors;
                            obj.GetType().InvokeMember("ElementsToRecordset", BindingFlags.InvokeMethod, null, obj, arg);
                        }
                        catch (Exception e)
                        {
                            session.mo_errobj = e;
                            engine.error("18 delete");
                        }
                        try
                        {
                            arg[0] = ors;
                            wrapper.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, wrapper, arg);
                        }
                        catch (Exception) { }
                        break;

                    //---------------------------- INSERT ---------------------------------------
                    case enInstrType.INSERT:    // appel de la méthode Insert
                        wrapper = engine.wrappers["WR_" + instruction.getparam(0)];		// nom wrapper
                        ors = wrapper.GetType().InvokeMember("RecordStructure", BindingFlags.GetProperty, null, wrapper, null);
                        obj = null;
                        try
                        {
                            xmls = engine.XmlServiceFactory2();
                            obj = xmls.GetType().InvokeMember("Service", BindingFlags.GetProperty, null, wrapper, null);
                            arg[0] = ors;
                            obj.GetType().InvokeMember("ElementsToRecordset", BindingFlags.InvokeMethod, null, obj, arg);
                        }
                        catch (Exception e)
                        {
                            session.mo_errobj = e;
                            engine.error("19 insert");
                        }
                        try
                        {
                            arg[0] = ors;
                            wrapper.GetType().InvokeMember("Cree", BindingFlags.InvokeMethod, null, wrapper, arg);
                        }
                        catch (Exception) { }
                        try
                        {
                            arg[0] = ors;
                            obj.GetType().InvokeMember("RecordsetToElements", BindingFlags.InvokeMethod, null, obj, arg);
                        }
                        catch (Exception e)
                        {
                            session.mo_errobj = e;
                            engine.error("20 insert");
                        }
                        break;
                    //--------------------------------- NULL -----------------------------------
                    case enInstrType.NULL:
                        break;
        
                    //---------------------------------------------------------------------------------------------
                    //---------------------------------------------- CODES RETOUR EGALEMENT TRAITES COMME ORDRES
                    //---------------------------------------------------------------------------------------------
                    case enInstrType.ERROR:
                        if (ptrlastaff < 0)
                        {
                          session.mo_return="ERROR";
                          return;		// retour script supérieur
                        }
                        ptr = ptrlastaff;
                        stack = (Stack)laststack.Clone(); 
                        break;

                    default:
                        break;
                }
                mo_skipcase = false;
                fininstr(ref mo_skipcase);
            }
            engine.ctx = tmpctx;
            mo_skipcase=false;// pour CASE+skip on ne passe pas par là
            mo_rc="";

            if (( c = session.mo_request)!= "")			// retour Server.exec
            {
                nctx = new Context(engine, (new Transaction("request", c)).compile(), "request");
                nctx.btransfer=true;
                nctx.brequest=brequest;
                session.mo_ctx=nctx;  // le nouveau contexte est conservé dans la session				
                session.mo_return=""; 
                session.mo_request = "";

                nctx.play();  // on l'exécute en récursif .   Au cas où l'on revient :

                if (nctx.brequest && nctx.btransfer)
                    engine.callstack.Pop();
                session.mo_ctx = (Context)callstack.Pop();
            }
        }

        private void fininstr(ref bool pmo_skipcase)
        {
            // fin d'instruction (;)
            bool bcont = true;
            Ctxinstr ct;

            while (bcont && stack.Count > 0)
            {
                ct = (Ctxinstr)stack.Pop();
                switch (ct.verb)
                {
                    case enInstrType.BLOCK:
                        stack.Push(ct);
                        bcont = false;
                        break;

                    case enInstrType.CASE:
                        pmo_skipcase = true;
                        break;

                    case enInstrType.REPEAT:
                        ptr = (int)ct.data;
                        pmo_skipcase = false;
                        bcont = false;
                        break;

                    case enInstrType.CALL:
                        pmo_skipcase = false;
                        break;
                }
            }
        }
    }

    private struct Ctxinstr
    {
        public enInstrType verb;
        public object data;
        public Ctxinstr(enInstrType v, object d)
        {
            verb = v;
            data = d;
        }
    }

    public Transactions transactions;

    public Engine(Appdata pdata)  
    {
        Application = pdata;
        site = pdata.mo_site;
        transactions = Application.script;

        session = new UnsavedSessionData();
        session.mo_intrans = false;
        session.mo_messlu = false;
        session.mo_menu = "NASTER";
        callstack = new Stack();
        savestack = new Stack();
        connections = new Hashtable();
        wrappers = new Hashtable();
        rform = new Hashtable();

        session.mo_xsl = new MSXML2.XSLTemplate();
        session.mo_xsl2 = new MSXML2.XSLTemplate();
        MSXML2.FreeThreadedDOMDocument tmpxmldom = new MSXML2.FreeThreadedDOMDocument();
        tmpxmldom.async = false;
        string nxsl = HttpContext.Current.Server.MapPath("./xsl/SERVICE-1.0-TO-XHTML-1.0-FRAGMENT.XSL");
        bool r = tmpxmldom.load(nxsl);
        if (!r)
            error("2 échec load XSL: " + nxsl);
        session.mo_xsl.stylesheet = tmpxmldom;
        tmpxmldom = null;
        menutransaction = transactions.get("MENU");
        session.mo_return = "";
        ////////////////////////////////////
        ssession.mo_nosave = false;
        session.mo_errmess = "";
        session.mo_nompage = "";
        session.mo_request = "";
        session.mo_intrans = false;
        session.mo_langue = "undefined";
        session.mo_clear = false;
        ssession.mo_lastreturn = "";
        session.mo_application = "NASTER";
        session.mo_caller = "";
        session.ex_utilisateur = "";
        session.mo_haserror = false;
        session.mo_noalerterror = false;
        ssession.mo_old_menu = "";
        string tmp = K.Request.ServerVariables["SERVER_NAME"];
        int i = tmp.IndexOf(".");
        if (i >= 0)
            tmp = tmp.Substring(0, i);
        //tmp = tmp.Substring(0, 10);
        session.mo_server = tmp;
        if (Application.mo_server == null)
            Application.mo_server = tmp;
        session.mo_viaxml = 0;      // 0=sans 1=transform 2=send avec xml
        session.mo_messlu = false;
        session.mo_messext = true;
        session.mo_startmenu = "TRANS MENU;";
        session.mo_clear = false;
        session.mo_modal = "NO";
        session.mo_alert = false;
        ////////////////////////
        K.Session["MO_RESULT"]= "";
        ssession.mo_commandes = "";
        session.mo_alert = false;
        theEngine = this;
    }
    //public object execute(Instruction pinstr)
    //{
    //}

    public object getwrapper(string pclass)
    {
       // object x = new plWrapper.act();
       // Assembly[] ass=AppDomain.CurrentDomain.GetAssemblies();// "Interop.plWrapper"ass[15].FullName
        object o = AppDomain.CurrentDomain.CreateInstanceAndUnwrap("Interop.plWrapper", "plWrapper.stylClass");//System.Reflection.Assembly.FullName
       // object o = Activator.CreateComInstanceFrom(@"C:\Inetpub\wwwroot\moniteurweb\Bin\Interop.plWrapper.dll", "plWrapper.stylClass").Unwrap(); 

        /*
        object[] arg = new object[1];
        plWrapper.glo w = new plWrapper.glo();
        object obj = conn;
        w.set_Connection(ref obj);
         */
        return o;
    }

    public void setstart(string pname)
    {
        transaction = transactions.get(pname);
        ctx = new Context(this,transaction, "");
    }
    public object run(string pname)
    {
        setstart(pname);
        return null;
    }
    public object run()
    {
        resume("trans naster;", false);
        return null;
    }
    public void resume(string preq,bool pmodal)////////////////////////////////////////////////////////////////
    {
        init();
        string req = preq;
        bool bmodal = pmodal;
        //HttpSessionState Session = HttpContext.Current.Session;
        bool bForm = (K.Request.ServerVariables["REQUEST_METHOD"] == "POST"); // encore mieux
        xxxtran = 0;
        /*
        ssession.mo_nosave  = false;

        session.mo_errmess = "";
        session.mo_nompage = "";
        session.mo_request = "";
        session.mo_intrans = false;
        session.mo_langue  = "undefined";
        session.mo_clear   = false;
        ssession.mo_lastreturn = "";
        session.mo_application = "NASTER";
        session.mo_caller  = "";
        session.ex_utilisateur = "";
        session.mo_haserror = false;
        session.mo_noalerterror = false;
        ssession.mo_old_menu = "";
        string tmp = K.Request.ServerVariables["SERVER_NAME"];
        int i = tmp.IndexOf(".");
        if (i >= 0)
            tmp = tmp.Substring(0, i);
        //tmp = tmp.Substring(0, 10);
        session.mo_server = tmp;
        if (Application.mo_server == null)
            Application.mo_server = tmp;
        session.mo_viaxml = 0;      // 0=sans 1=transform 2=send avec xml
        session.mo_messlu	=false;
        session.mo_messext = true;
        session.mo_startmenu = "TRANS MENU;";
        session.mo_clear = false;
        session.mo_modal = "NO";
        session.mo_alert = false;
        */

        while (true)														// boucle sur retour au menu départ
        {
            session.mo_save_callstack = callstack.Count;
            session.mo_save_savestack = savestack.Count;

            if (req != null && req != "")
            {    
                // pour une fenetre modale on sauve le niveau des piles en cas de sortie anormale (fermeure par bouton systeme)
                if (bmodal)
                {
                    K.Session["MO_SAVE_CALLSTACK"] = session.mo_save_callstack;
                    K.Session["MO_SAVE_SAVESTACK"] = session.mo_save_savestack;
                    bmodal = false;
                }
                // sauvegarde script en cours si nécessaire
                if (session.mo_ctx != null)
                {
                    ctx = session.mo_ctx;
                    callstack.Push(ctx); 
                }
                // initialisation script
                ctx = new Context(this,(new Transaction("request", req)).compile(), "request");
                ctx.brequest = true;
                session.mo_ctx = ctx;
            }
            req = null;

            //boucle sur les scripts imbriqués

            while (true)				// boucle sur transactions empilées
            {
                session.mo_ctx.play();
                if (callstack.Count == 0)
                    break;
                if (ctx.brequest && ctx.btransfer)
                {
                    if (callstack.Count == 0)
                        break;
                    callstack.Pop();
                }
                session.mo_ctx =ctx= (Context)callstack.Pop();
                if (ctx.bissaved)
                {
                    restore();
                    session.mo_ctx.bissaved = false;
                }
                session.mo_transaction = ctx.transaction.name;
                session.mo_caller = ctx.caller;
            }
            session.mo_ctx = new Context(this,menutransaction , "menu");
        }
    }
    private void restore()
    {
        session.mo_ctx.restore();
    }
    public void compile()
    {
        transactions.compile();
    }
    public string getpath(string f)
    {
	    if(f.EndsWith(".XML"))
		    return "/"+site+"/xml/"+f;
        if (f.StartsWith("MDL_"))
            return f.ToLower();
	    if(f.IndexOf(".")<0)
	      f+=".ASPX";

        switch (f.Substring(0, 4))
	    {
		    case "PGL_": return "/"+site+"/asp/v/pgl/"+f; 
		    case "TPL_": return "/"+site+"/asp/v/tpl/"+f; 
	    }
	    return "/"+site+"/"+f;
    }
    public string gloinit(ADODB.Connection conn)
    {
        object[] arg=new object[1];
        plWrapper.glo w = new plWrapper.glo();
        object obj = conn;
        w.Connection = conn;
        // initialisation des variables globales de GLO
        w.Initialise();
        // chargement des variables
        //object rs=w.Charge();
        ADODB.Recordset rs=w.Charge();
        // modification
        string l=langue;
        l=(l=="undefined")?"*":l;
        rs.Fields["co_langue"].Value = l;
        rs.Fields["id_utilisateur"].Value = Agent.theAgent.mo_user;
        /*
        object fields = rs.GetType().InvokeMember("Fields", BindingFlags.GetField, null, rs, null);
        arg[0] = "co_langue";
        object field = fields.GetType().InvokeMember("Item", BindingFlags.GetField, null, fields, arg);
        arg[0] = l;
        field.GetType().InvokeMember("Value", BindingFlags.GetField, null, field, arg);
        arg[0] = "id_utilisateur";
        field = fields.GetType().InvokeMember("Item", BindingFlags.GetField, null, fields, arg);
        arg[0] = Agent.theAgent.mo_user;
        field.GetType().InvokeMember("Value", BindingFlags.GetField, null, field, arg);
        */
        obj = rs;
        w.Modifie(ref obj);

        plWrapper.err_Renamed w2 = new plWrapper.err_Renamed();
        obj = conn;
        w2.Connection=conn;
        obj = session.mo_application;
        w2.Enregistreapplication(ref obj);
        obj = session.mo_transaction;
        w2.Enregistretransaction(ref obj,ref obj);
        return Convert.ToString(w.SessionID);
    }
    public static void logmess(string src, string mess)
    {
        try
        {
            HttpContext.Current.Application.Lock();
            StreamWriter sw = new StreamWriter("c:\\winnt\\system32\\logfiles\\naster\\errors.txt", true);
            sw.Write("*****logmess= " + theEngine.session.ex_utilisateur + " " + DateTime.Now.ToString("g") + " (" + src + ") " + mess + "\r\n");
            sw.Close();
        }
        catch (Exception) { }
        finally 
        {
            HttpContext.Current.Application.UnLock();
        }
    }
    public string getAspSessionID()
    {
        string h = K.Request.ServerVariables["HTTP_COOKIE"];
        if (h==null)
            return null;
        int i = h.IndexOf("ASPSESSIONID");
        if (i < 0)
            return "";
        int k = h.IndexOf(";", i);
        if (k < 0)
            return h.Substring(i);
        else
            return h.Substring(i, k-i+1);
    }
    public static void mailtogc(string mess)
    {
        SmtpClient cli = new SmtpClient();
        cli.Send("naster@mcsfr.com", "bugs@mcsfr.com", "[gcNASTER]ESPION NASTER utilisateur="+theEngine.session.ex_utilisateur+" serveur="+theEngine.session.mo_server, mess);
    }
    private string stname,stinstr;
    private int stline;
    public string descstack()
    {
        string fic,r="\r\n";
        Stack s;//,si,ctx;

        fic=r+r+"-------------------------------------------------------"+r;
        fic+=DateTime.Now.ToString("g")  +" Utilisateur: "+session.ex_utilisateur+r+r;
        fic+="PILE D'APPELS:"+r+r;
        s=callstack;
        if(s!=null)
        {
            foreach(Engine.Context si in s)
            { 
              getpos(si);
              fic+="transaction: "+si.transaction.name+" ligne: "+System.Convert.ToString(stline)+" "+stinstr+r;
            }
        }
        getpos(session.mo_ctx);
        fic += "transaction: " + stname + " ligne: " + System.Convert.ToString(stline) + " " + stinstr + r;
        return fic;
    }

    public void getpos(Engine.Context ctx)
    {
         stname="";
         stinstr="";
         stline=0;
         
         if(ctx==null)
           return;

         stname=ctx.transaction.name;

         stline=((Instruction)(ctx.transaction.instructions[ctx.ptr - 1])).linenum;
    }

    public static string getvalue(string strname)
    {
        object o;
        if (theEngine.session.mo_viaxml == 2) // n utiliser XMLS que s'il y a eu transform et pas NOSAVE 
        {
            try
            {
                return (string)theEngine.XmlServiceFactory2().service.elementlist[strname].SingleValue;
            }
            catch (Exception)
            {
                return (string)theEngine.rform[strname.ToUpper()];
            }
        }
        else
            return (string)theEngine.rform[strname.ToUpper()];
    }
}

public class Transaction // source de module/transaction
{
    public string name;
    public string prg;
    public ArrayList instructions;
    public ArrayList strings;
    public ArrayList symbols;
    private Hashtable htable;
    public Transaction(string pname, string pprg)
    {
        name = pname;
        prg = pprg;
    }
    public int addsymbole(string s)
    {
        if (htable.Contains(s))
            return (int)htable[s];
        int n = strings.Add(s);
        htable[s] = n;
        return n;
    }
    public int addstring(string s)
    {
        return addsymbole(s);
    }
    private ArrayList lines=new ArrayList();
    private void setlines()
    {
        int i=0;
        do
        {
            i=prg.IndexOf("\n",i);
            if (i >= 0)
            {
                lines.Add(i);
                i += 2;
            }

        }
        while(i>=0 && i<prg.Length);
    }
    private string text;
    public short getline(int ppos)
    {
        int p1,p2;
        for (short i = 0; i < lines.Count; i++)
            if (ppos < (p2=(int)lines[i]))
            {
                if (i == 0)
                    p1 = 0;
                else
                    p1 = (int)lines[i - 1]+2;
                text = prg.Substring(p1, p2 - p1 + 1);
                return (short)(i + 1);
            }
        if (lines.Count == 0)
            text = prg;
        else
        {
            p1 = (int)lines[lines.Count - 1]+2;
            if (p1 < prg.Length)
                text = prg.Substring(p1);
            else
                text = "";
        }
        return (short)(lines.Count+1);
    }

    public Transaction compile()
    {
        setlines();

        instructions = new ArrayList();
        symbols = new ArrayList();
        strings = new ArrayList();
        htable = new Hashtable();

        int p = 0;
        string verb, token = "", s;
        Instruction instr;
        InstructionType instrtype;
        enInstrType entype;
        int y;

        while (true)
        {
            if (token == ";" || token == "")
                verb = Parser.getToken(prg, ref p);
            else
                verb = token;
            if (verb == null)
                break;

            entype = Language.getType(verb);
            instr = new Instruction(this,entype);
            instr.linenum = getline(p);

            instrtype = Language.getInstructionType(entype);
            token = "";

            if (instrtype.parms != null)
                foreach (ParamType parm in instrtype.parms)
                {
                    while (true)
                    {
                        if (entype != enInstrType.COMMENT && entype != enInstrType.EVAL2)
                        {
                            if (token == "")
                                token = Parser.getToken(prg, ref p);
                            if (token == ";")
                                break;
                        }
                        else
                            token = "";

                        switch (parm.type)
                        {
                            case enParamType.KEYWORD:
                                if (Parser.isword(token))
                                    for (int i = 0; i < parm.keywords.Length; i++)
                                        if ((s = parm.keywords[i]) == token)
                                        {
                                            instr.bkeywords[i] = true;
                                            token = "";
                                            break;
                                        }
                                break;

                            case enParamType.SYMBOL:
                                if (Parser.isword(token))
                                {
                                    y=addsymbole(token);
                                    instr.parms.Add(y);
                                    token = "";
                                }
                                break;

                            case enParamType.STRING:
                                if (token == "§")
                                {
                                    p--;
                                    s = Parser.getEval(prg, ref p);
                                    y = strings.Add(s);
                                    instr.parms.Add(-y);
                                    token = "";
                                    break;
                                }
                                if (token == "\"")
                                {
                                    p--;
                                    s = Parser.getString(prg, ref p);
                                    y = addstring(s);
                                    instr.parms.Add(y);
                                    token = "";
                                }
                                break;

                            case enParamType.EVAL:
                                if (token == "§" || token == "")
                                {
                                    p--;
                                    s = Parser.getEval(prg, ref p);
                                    y = strings.Add(s);
                                    instr.parms.Add(y);
                                    token = "";
                                }
                                break;

                            case enParamType.COMMENT:
                                p--;
                                s = Parser.getComment(prg, ref p);
                                instr = null;
                                break;

                            default:
                                break;
                        }
                        if (!parm.bmultiple)
                            break;
                    }
                    if (token == ";")
                        break;
                }
            if (instr != null)
                instructions.Add(instr);
            token = Parser.getToken(prg, ref p);
        }
        //prg = null;
        return this;
    }
}
public class Transactions // ensemble des transactions
{
    private Hashtable hash = new Hashtable();
    private Transaction[] transactions;
    private int n;

    public Transactions(int pmax)
    {
        transactions = new Transaction[pmax];
        n = 0;
    }
    public void put(Transaction ptrans)
    {
        hash[ptrans.name.ToUpper()] = n;
        transactions[n] = ptrans;
        n++;
    }
    public Transaction get(string pname)
    {
        return transactions[(int)hash[pname.ToUpper()]];
    }
    public string getprg(string pname)
    {
        return get(pname).prg;
    }
    public int getposition(string pname)
    {
        return (int)hash[pname.ToUpper()];
    }
    public void compile()
    {
        foreach (Transaction t in transactions)
            if (t != null)
                t.compile();
    }

    public static Transactions readscript(string ppath)
    {
        Regex r1 = new Regex("deftrans ", RegexOptions.IgnoreCase);
        Regex r2 = new Regex("enddef", RegexOptions.IgnoreCase);
        Match m;
        Encoding enc = Encoding.GetEncoding(1252);  // Windows default Code Page
        FileStream fs = new FileStream(ppath, FileMode.Open);
        StreamReader sr = new StreamReader(fs, enc);
        string text = sr.ReadToEnd();
        int p = 0;
        string name, source;
        Transaction transaction;

        Transactions transactions = new Transactions(500);
        while (true)
        {
            m = r1.Match(text, p);
            if (!m.Success)
                break;
            p = m.Index + 9;
            name = Parser.getToken(text, ref p);
            Parser.getToken(text, ref p); //;
            m = r2.Match(text, p);
            source = text.Substring(p, m.Index - p);
            transaction = new Transaction(name, source);
            transactions.put(transaction);
        }
        sr.Close();
        fs.Close();
        return transactions;
    }
}

public class InstructionType  // type d'instruction
{
    public enInstrType type;
    public string verb;
    public ParamType[] parms;

    public InstructionType(enInstrType ptype, string pverb, ParamType[] pparms)
    {
        type = ptype;
        verb = pverb;
        parms = pparms;
    }
}
public struct ParamType  // type de parametre
{
    public enParamType type;
    public bool boptional;
    public bool bmultiple;
    public string[] keywords;

    public ParamType(enParamType ptype, bool pboptional, bool pbmultiple, string[] pkeywords)
    {
        type = ptype;
        boptional = pboptional;
        bmultiple = pbmultiple;
        keywords = pkeywords;
    }
}
public class Instruction  // instruction de transaction
{
    public Transaction transaction;
    public enInstrType type;
    public bool[] bkeywords;
    public ArrayList parms;
    public short linenum;
    public string text;
    public Instruction(Transaction t, enInstrType ptype)
    {
        transaction = t;
        type = ptype;
        bkeywords = new bool[5];
        parms = new ArrayList();
    }
    public bool haskeyword(int n)
    {
        return bkeywords[n];
    }
    public string getparam(int n)
    {
        if (n >= parms.Count)
            return null;
        return (string)transaction.strings[(int)parms[n]];
    }
    public ArrayList getlist(int n)
    {
        ArrayList a = new ArrayList();
        for(int i=n;i<parms.Count;i++)
            a.Add(getparam(i));
        return a;
    }
}

//public struct ParamValue  // valeur de parametre
//{
//}

public class Language
{
    //private static string whitespaces = " \t\n";
    private static InstructionType[] instructions;
    static Language()
    {
        ParamType[] lparms;

        instructions = new InstructionType[64];

        instructions[0] = new InstructionType(enInstrType.ABANDON, "ABANDON", null);
        instructions[1] = new InstructionType(enInstrType.ABORT, "ABORT", null);
        instructions[2] = new InstructionType(enInstrType.RETURN, "RETURN", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.STRING, false, false, null);
        instructions[3] = new InstructionType(enInstrType.ALERT, "ALERT", lparms);

        lparms = new ParamType[2];
        lparms[0] = new ParamType(enParamType.STRING, false, false, null);
        lparms[1] = new ParamType(enParamType.SYMBOL, true, false, null);
        instructions[4] = new InstructionType(enInstrType.TITLE, "TITLE", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[5] = new InstructionType(enInstrType.WRNAME, "WRNAME", lparms);

        instructions[6] = new InstructionType(enInstrType.BP, "BP", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.KEYWORD, false, false, new string[3] { "ON", "START", "END" });
        instructions[7] = new InstructionType(enInstrType.TLOG, "TLOG", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[8] = new InstructionType(enInstrType.TLOGBEG, "TLOGBEG", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[9] = new InstructionType(enInstrType.TLOGEND, "TLOGEND", lparms);

        lparms = new ParamType[2];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        lparms[1] = new ParamType(enParamType.KEYWORD, false, false, new string[2] { "TRUE", "FALSE" });
        instructions[10] = new InstructionType(enInstrType.SQLTRACE, "SQLTRACE", lparms);


        lparms = new ParamType[3];
        lparms[0] = new ParamType(enParamType.KEYWORD, true, false, new string[1] { "=" });
        lparms[1] = new ParamType(enParamType.STRING, false, false, null);
        lparms[2] = new ParamType(enParamType.SYMBOL, true, false, null);
        instructions[11] = new InstructionType(enInstrType.COMMANDES, "COMMANDES", lparms);

        instructions[12] = new InstructionType(enInstrType.NOCOMMANDES, "NOCOMMANDES", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[13] = new InstructionType(enInstrType.LOADXML, "LOADXML", lparms);

        instructions[14] = new InstructionType(enInstrType.SWAPXML, "SWAPXML", null);
        instructions[15] = new InstructionType(enInstrType.CLOSEXML, "CLOSEXML", null);

        lparms = new ParamType[2];
        lparms[0] = new ParamType(enParamType.SYMBOL, true, false, null);
        lparms[1] = new ParamType(enParamType.SYMBOL, true, false, null);
        instructions[16] = new InstructionType(enInstrType.TRANSFORM, "TRANSFORM", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[17] = new InstructionType(enInstrType.SENDDATA, "SENDDATA", lparms);

        instructions[18] = new InstructionType(enInstrType.DISP, "DISP", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[19] = new InstructionType(enInstrType.RAW, "RAW", lparms);

        lparms = new ParamType[2];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        lparms[1] = new ParamType(enParamType.SYMBOL, true, false, null);
        instructions[20] = new InstructionType(enInstrType.VIEW, "VIEW", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[21] = new InstructionType(enInstrType.EXECP, "EXECP", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[22] = new InstructionType(enInstrType.EXEC, "EXEC", lparms);

        instructions[23] = new InstructionType(enInstrType.IMBRICATE, "IMBRICATE", null); // SUPPRIME
        instructions[24] = new InstructionType(enInstrType.CALL, "CALL", null);
        instructions[25] = new InstructionType(enInstrType.SAVE, "SAVE", null);
        instructions[26] = new InstructionType(enInstrType.RESTORE, "RESTORE", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[27] = new InstructionType(enInstrType.TRANS, "TRANS", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[28] = new InstructionType(enInstrType.BRANCH, "BRANCH", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[29] = new InstructionType(enInstrType.TRANSFER, "TRANSFER", lparms);

        lparms = new ParamType[2];
        lparms[0] = new ParamType(enParamType.KEYWORD, true, false, new string[1] { "NOT" });
        lparms[1] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[30] = new InstructionType(enInstrType.CASE, "CASE", lparms);

        instructions[31] = new InstructionType(enInstrType.OTHER, "OTHER", null);
        instructions[32] = new InstructionType(enInstrType.REPEAT, "REPEAT", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[33] = new InstructionType(enInstrType.ON, "ON", lparms);

        instructions[34] = new InstructionType(enInstrType.CONTINUE, "CONTINUE", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[35] = new InstructionType(enInstrType.LABEL, "LABEL", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.EVAL, false, false, null);
        instructions[36] = new InstructionType(enInstrType.EVAL, "EVAL", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[37] = new InstructionType(enInstrType.JUMP, "JUMP", lparms);

        instructions[38] = new InstructionType(enInstrType.JUMPPREV, "JUMPPREV", null);
        instructions[39] = new InstructionType(enInstrType.JUMPNEXT, "JUMPNEXT", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[38] = new InstructionType(enInstrType.READTRANS, "READTRANS", lparms);

        lparms = new ParamType[2];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        lparms[1] = new ParamType(enParamType.SYMBOL, true, true, null);
        instructions[39] = new InstructionType(enInstrType.GETCONNECTION, "GETCONNECTION", lparms);

        lparms = new ParamType[2];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        lparms[1] = new ParamType(enParamType.SYMBOL, true, true, null);
        instructions[40] = new InstructionType(enInstrType.GETCONNECTION2, "GETCONNECTION2", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[41] = new InstructionType(enInstrType.GLOINIT, "GLOINIT", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[42] = new InstructionType(enInstrType.CLOSECONNECTION, "CLOSECONNECTION", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[43] = new InstructionType(enInstrType.CLOSECONNECTION2, "CLOSECONNECTION2", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, true, true, null);
        instructions[44] = new InstructionType(enInstrType.BEGINTRANS, "BEGINTRANS", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, true, true, null);
        instructions[45] = new InstructionType(enInstrType.COMMIT, "COMMIT", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, true, true, null);
        instructions[46] = new InstructionType(enInstrType.ROLLBACK, "ROLLBACK", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, true, null);
        instructions[47] = new InstructionType(enInstrType.GETWRAPPER, "GETWRAPPER", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, true, null);
        instructions[48] = new InstructionType(enInstrType.CLOSEWRAPPER, "CLOSEWRAPPER", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[49] = new InstructionType(enInstrType.UPDATE, "UPDATE", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[50] = new InstructionType(enInstrType.DELETE, "DELETE", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        instructions[51] = new InstructionType(enInstrType.INSERT, "INSERT", lparms);

        instructions[52] = new InstructionType(enInstrType.LOOP, "LOOP", null);
        instructions[53] = new InstructionType(enInstrType.BREAK, "BREAK", null);
        instructions[54] = new InstructionType(enInstrType.ENDMODAL, "ENDMODAL", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.COMMENT, true, false, null);
        instructions[55] = new InstructionType(enInstrType.COMMENT, "COMMENT", lparms);

        instructions[56] = new InstructionType(enInstrType.BLOCK, "(", null);
        instructions[57] = new InstructionType(enInstrType.ENDBLOCK, ")", null);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.EVAL, false, false, null);
        instructions[58] = new InstructionType(enInstrType.EVAL2, "§", lparms);

        lparms = new ParamType[3];
        lparms[0] = new ParamType(enParamType.SYMBOL, false, false, null);
        lparms[1] = new ParamType(enParamType.KEYWORD, false, false, new string[1] { "=" });
        lparms[2] = new ParamType(enParamType.STRING, true, false, null);
        instructions[59] = new InstructionType(enInstrType.DEFINEDATASOURCE, "DEFINEDATASOURCE", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.KEYWORD, true, true, new string[] { "NOSAVE", "CLEAR", "NORETURN", "NOCONTENT" });
        instructions[60] = new InstructionType(enInstrType.SEND, "SEND", lparms);

        lparms = new ParamType[1];
        lparms[0] = new ParamType(enParamType.SYMBOL, true, false, null);
        instructions[61] = new InstructionType(enInstrType.LANGUAGE, "LANGUAGE", lparms);

        instructions[62] = new InstructionType(enInstrType.BLOCK, "ERROR", null);

        instructions[63] = new InstructionType(enInstrType.NULL, "NULL", null);
    }
    public static InstructionType getInstructionType(string pverb)
    {
        enInstrType ptype = Language.getType(pverb);
        return getInstructionType(ptype);
    }
    public static InstructionType getInstructionType(enInstrType ptype)
    {
        foreach (InstructionType p in instructions)
            if (p.type == ptype)
                return p;
        return null;
    }
    public static Instruction getInstruction(string psource, ref int pptr)
    {
        return (Instruction)null;
    }
    public static enInstrType getType(string pverb)
    {
        switch (pverb)
        {
            case "DEFINEDATASOURCE": return enInstrType.DEFINEDATASOURCE;
            case "ABANDON": return enInstrType.ABANDON;
            case "ABORT": return enInstrType.ABORT;
            case "RETURN": return enInstrType.RETURN;
            case "ALERT": return enInstrType.ALERT;
            case "TITLE": return enInstrType.TITLE;
            case "WRNAME": return enInstrType.WRNAME;
            case "BP": return enInstrType.BP;
            case "TLOG": return enInstrType.TLOG;
            case "TLOGBEG": return enInstrType.TLOGBEG;
            case "TLOGEND": return enInstrType.TLOGEND;
            case "SQLTRACE": return enInstrType.SQLTRACE;
            case "COMMANDES": return enInstrType.COMMANDES;
            case "NOCOMMANDES": return enInstrType.NOCOMMANDES;
            case "LOADXML": return enInstrType.LOADXML;
            case "SWAPXML": return enInstrType.SWAPXML;
            case "CLOSEXML": return enInstrType.CLOSEXML;
            case "TRANSFORM": return enInstrType.TRANSFORM;
            case "SENDDATA": return enInstrType.SENDDATA;
            case "DISP": return enInstrType.DISP;
            case "RAW": return enInstrType.RAW;
            case "VIEW": return enInstrType.VIEW;
            case "SEND": return enInstrType.SEND;
            case "EXEC": return enInstrType.EXEC;
            case "EXECP": return enInstrType.EXECP;
            case "IMBRICATE": return enInstrType.IMBRICATE;// SUPPRIME
            case "CALL": return enInstrType.CALL;
            case "SAVE": return enInstrType.SAVE;
            case "RESTORE": return enInstrType.RESTORE;
            case "TRANS": return enInstrType.TRANS;
            case "BRANCH": return enInstrType.BRANCH;
            case "TRANSFER": return enInstrType.TRANSFER;
            case "CASE": return enInstrType.CASE;
            case "OTHER": return enInstrType.OTHER;
            case "REPEAT": return enInstrType.REPEAT;
            case "ON": return enInstrType.ON;
            case "CONTINUE": return enInstrType.CONTINUE;
            case "LABEL": return enInstrType.LABEL;
            case "EVAL": return enInstrType.EVAL;
            case "JUMP": return enInstrType.JUMP;
            case "JUMPPREV": return enInstrType.JUMPPREV;
            case "JUMPNEXT": return enInstrType.JUMPNEXT;
            case "READTRANS": return enInstrType.READTRANS;
            case "GETCONNECTION": return enInstrType.GETCONNECTION;
            case "GETCONNECTION2": return enInstrType.GETCONNECTION2;
            case "GLOINIT": return enInstrType.GLOINIT;
            case "CLOSECONNECTION": return enInstrType.CLOSECONNECTION;
            case "CLOSECONNECTION2": return enInstrType.CLOSECONNECTION2;
            case "BEGINTRANS": return enInstrType.BEGINTRANS;
            case "COMMIT": return enInstrType.COMMIT;
            case "ROLLBACK": return enInstrType.ROLLBACK;
            case "GETWRAPPER": return enInstrType.GETWRAPPER;
            case "CLOSEWRAPPER": return enInstrType.CLOSEWRAPPER;
            case "UPDATE": return enInstrType.UPDATE;
            case "DELETE": return enInstrType.DELETE;
            case "INSERT": return enInstrType.INSERT;
            case "LOOP": return enInstrType.LOOP;
            case "BREAK": return enInstrType.BREAK;
            case "ENDMODAL": return enInstrType.ENDMODAL;
            case "LANGUAGE": return enInstrType.LANGUAGE;
            case "*": return enInstrType.COMMENT;
            case "(": return enInstrType.BLOCK;
            case "§": return enInstrType.EVAL2;
            case ")": return enInstrType.ENDBLOCK;
            case "ERROR": return enInstrType.ERROR;
            case "NULL": return enInstrType.NULL;
            default:
                Exception e = new Exception("verbe inconnu : " + pverb);
                throw e;
        }
    }
    public static string getVerb(enInstrType ptype)
    {
        switch (ptype)
        {
            case enInstrType.DEFINEDATASOURCE: return "DEFINEDATASOURCE";
            case enInstrType.ABANDON: return "ABANDON";
            case enInstrType.ABORT: return "ABORT";
            case enInstrType.RETURN: return "RETURN";
            case enInstrType.ALERT: return "ALERT";
            case enInstrType.TITLE: return "TITLE";
            case enInstrType.WRNAME: return "WRNAME";
            case enInstrType.BP: return "BP";
            case enInstrType.TLOG: return "TLOG";
            case enInstrType.TLOGBEG: return "TLOGBEG";
            case enInstrType.TLOGEND: return "TLOGEND";
            case enInstrType.SQLTRACE: return "SQLTRACE";
            case enInstrType.COMMANDES: return "COMMANDES";
            case enInstrType.NOCOMMANDES: return "NOCOMMANDES";
            case enInstrType.LOADXML: return "LOADXML";
            case enInstrType.SWAPXML: return "SWAPXML";
            case enInstrType.CLOSEXML: return "CLOSEXML";
            case enInstrType.TRANSFORM: return "TRANSFORM";
            case enInstrType.SENDDATA: return "SENDDATA";
            case enInstrType.DISP: return "DISP";
            case enInstrType.RAW: return "RAW";
            case enInstrType.VIEW: return "VIEW";
            case enInstrType.SEND: return "SEND";
            case enInstrType.EXEC: return "EXEC";
            case enInstrType.EXECP: return "EXECP";
            case enInstrType.IMBRICATE: return "IMBRICATE";// SUPPRIME
            case enInstrType.CALL: return "CALL";
            case enInstrType.SAVE: return "SAVE";
            case enInstrType.RESTORE: return "RESTORE";
            case enInstrType.TRANS: return "TRANS";
            case enInstrType.BRANCH: return "BRANCH";
            case enInstrType.TRANSFER: return "TRANSFER";
            case enInstrType.CASE: return "CASE";
            case enInstrType.OTHER: return "OTHER";
            case enInstrType.REPEAT: return "REPEAT";
            case enInstrType.ON: return "ON";
            case enInstrType.CONTINUE: return "CONTINUE";
            case enInstrType.LABEL: return "LABEL";
            case enInstrType.EVAL: return "EVAL";
            case enInstrType.JUMP: return "JUMP";
            case enInstrType.JUMPPREV: return "JUMPPREV";
            case enInstrType.JUMPNEXT: return "JUMPNEXT";
            case enInstrType.READTRANS: return "READTRANS";
            case enInstrType.GETCONNECTION: return "GETCONNECTION";
            case enInstrType.GETCONNECTION2: return "GETCONNECTION2";
            case enInstrType.GLOINIT: return "GLOINIT";
            case enInstrType.CLOSECONNECTION: return "CLOSECONNECTION";
            case enInstrType.CLOSECONNECTION2: return "CLOSECONNECTION2";
            case enInstrType.BEGINTRANS: return "BEGINTRANS";
            case enInstrType.COMMIT: return "COMMIT";
            case enInstrType.ROLLBACK: return "ROLLBACK";
            case enInstrType.GETWRAPPER: return "GETWRAPPER";
            case enInstrType.CLOSEWRAPPER: return "CLOSEWRAPPER";
            case enInstrType.UPDATE: return "UPDATE";
            case enInstrType.DELETE: return "DELETE";
            case enInstrType.INSERT: return "INSERT";
            case enInstrType.LOOP: return "LOOP";
            case enInstrType.BREAK: return "BREAK";
            case enInstrType.ENDMODAL: return "ENDMODAL";
            case enInstrType.LANGUAGE: return "LANGUAGE";
            case enInstrType.COMMENT: return "*";
            case enInstrType.BLOCK: return "(";
            case enInstrType.EVAL2: return "§";
            case enInstrType.ENDBLOCK: return ")";
            case enInstrType.ERROR: return "ERROR";
            case enInstrType.NULL: return "NULL";
            default:
                Exception e = new Exception("type inconnu : " + ptype);
                throw e;
        }
    }
}

public class Parser
{
    static Regex r1 = new Regex(@"[\(\);""§\=\*]|\b(\w+(\.\w+)?(:\w+)?)");
    static Regex r2 = new Regex(@"[\s]*(?:""([^§""]*)§[\s]*)*""([^§""]*)""");
    static Regex r3 = new Regex(@"[\s]*§([^§]*)§");
    static Regex r4 = new Regex(@"\*([^;]*);");
    static Match m;
    public static string getToken(string s, ref int p)
    {
        m = r1.Match(s, p);
        if (m.Success)
        {
            p = m.Index + m.Length;
            return m.Groups[0].Value.ToUpper();
        }
        else
            return null;
    }
    public static string getString(string s, ref int p)
    {

        m = r2.Match(s, p);
        string ret = "";
        foreach (Capture c in m.Groups[1].Captures)
            ret += c.Value;
        ret += m.Groups[2].Value;
        p = m.Index + m.Length;
        return ret;
    }
    public static string getEval(string s, ref int p)
    {

        m = r3.Match(s, p);
        string ret = m.Groups[1].Value;
        p = m.Index + m.Length;
        return ret;
    }
    public static string getComment(string s, ref int p)
    {

        m = r4.Match(s, p);
        string ret = m.Groups[1].Value;
        p = m.Index + m.Length;
        return ret;
    }
    public static bool isword(string s)
    {
        return s.Length > 1 || ";()\"§*".IndexOf(s) == -1;
    }
}



public enum enInstrType // instructions
{
    ABANDON, ABORT, RETURN, ALERT, TITLE, WRNAME, BP, TLOG, TLOGBEG, TLOGEND, SQLTRACE, COMMANDES, NOCOMMANDES,
    LOADXML, SWAPXML, CLOSEXML, TRANSFORM, SENDDATA, DISP, RAW, VIEW, SEND, EXEC, EXECP, IMBRICATE, CALL, SAVE, RESTORE,
    TRANS, BRANCH, TRANSFER, CASE, OTHER, REPEAT, ON, CONTINUE, LABEL, EVAL, JUMP, JUMPPREV, JUMPNEXT, READTRANS,
    GETCONNECTION, GETCONNECTION2, GLOINIT, CLOSECONNECTION, CLOSECONNECTION2, BEGINTRANS, COMMIT, ROLLBACK,
    GETWRAPPER, CLOSEWRAPPER, UPDATE, DELETE, INSERT, LOOP, BREAK, ENDMODAL,
    COMMENT, BLOCK, ENDBLOCK, EVAL2, DEFINEDATASOURCE,LANGUAGE,ERROR,NULL
}
public enum enParamType // types parametres
{
    STRING, SYMBOL, KEYWORD, PAGE, EVAL, COMMENT
}

public class Evaluator
{
    public static int EvalToInteger(string statement)
    {
        string s = EvalToString(statement);
        return int.Parse(s.ToString());
    }

    public static double EvalToDouble(string statement)
    {
        string s = EvalToString(statement);
        return double.Parse(s);
    }

    public static void EvalNoResult(string statement)
    {
        object o = EvalToObject(statement,false);
    }
    public static string EvalToString(string statement)
    {
        object o = EvalToObject(statement,true);
        return o.ToString();
    }
    
    public static object EvalToObject(string statement,bool bresult)
    {
        //public interface INaster; défini dans c:\classlib.dll
        // car cette méthode ET l'assembly générée doivent utiliser le meme type interface
        // et non pas deux types avec memes signatures

        INaster naster=new Naster();

        _jscriptSource = _jscriptSource1 + (bresult ? _jscriptSource1r : "") + statement + (bresult ? _jscriptSource2r : "") + _jscriptSource2;
        parameters.GenerateInMemory = true;
        parameters.OutputAssembly = @"c:\naster\assembly2.dll";
        results = prov.CompileAssemblyFromSource(parameters, _jscriptSource);
        if (results.Errors.HasErrors)
            throw new Exception("erreur Eval");
        assembly = results.CompiledAssembly;
        _evaluatorType = assembly.GetType("Evaluator.Evaluator");

        _evaluator = Activator.CreateInstance(_evaluatorType);

        object o= _evaluatorType.InvokeMember(
                    "Eval",
                    BindingFlags.InvokeMethod,
                    null,
                    _evaluator,
                    new object[] {naster} 
                 );
        _evaluator = null;
        _evaluatorType = null;
        assembly = null;
        results = null;
        return o;
    }
    static Evaluator()
    {
        prov = new CSharpCodeProvider();
        parameters = new CompilerParameters();
        parameters.GenerateInMemory = true;
        parameters.ReferencedAssemblies.Add("System.Web.dll");
        parameters.ReferencedAssemblies.Add(@"C:\naster\ClassLib.dll");
        _jscriptSource1 =
        @"  using System;
            using System.Web;
            using System.Web.SessionState;
            using ClassLib;
            namespace Evaluator
            {
               class Evaluator
               {
                  INaster m_naster;
                  public string Eval(object o) 
                  { 
                     m_naster=(INaster)o;
                     HttpSessionState Session  = HttpContext.Current.Session;
                     object result=null;";
        _jscriptSource1r="result=";
        _jscriptSource2r = ";";
        _jscriptSource2 =
                @"   return Convert.ToString(result);
                  }
                public string getvalue(string par)
                {
                    return m_naster.getvalue(par);
                }
                public int getdroit(string d)
                {
                    return m_naster.getDroit(d);
                }
                public void setdroit(string d, int pdroit)
                {
                    m_naster.setDroit(d,pdroit);
                }
                public string session(string s)
                {
                    return m_naster.getsession(s);
                }
                public int getuser()
                {
                    return m_naster.getuser();
                }
                public string getflavor()
                {
                    return m_naster.getflavor();
                }
               }
            }";
    }

    private static object _evaluator = null;
    private static Type _evaluatorType = null;
    private static CodeDomProvider prov;
    private static CompilerParameters parameters;
    private static CompilerResults results;
    private static Assembly assembly;
    private static string _jscriptSource1;
    private static string _jscriptSource1r;
    private static string _jscriptSource2;
    private static string _jscriptSource2r;
    private static string _jscriptSource;
}
public static class Utility
{
    public static void setcount(Stack vec,int n)       //modif taille caprockVec
    {
	    int c=vec.Count;
        while (c-- > n)
            vec.Pop();
    }
}
public class Naster : INaster
{
    public string getvalue(string par)
    {
        return Engine.getvalue(par);
    }
    public int getDroit(string d)
    {
        return Agent.theAgent.getDroit(d);
    }
    public void setDroit(string d, int pdroit)
    {
        Agent.theAgent.setDroit(d,pdroit);
    }
    public string getsession(string s)
    {
        switch (s)
        {
            case "MO_LASTRETURN":
                return Engine.theEngine.ssession.mo_lastreturn;
            default:
                return (string)K.Session[s];
        }
    }
    public int getuser()
    {
        return Agent.theAgent.mo_user;
    }
    public string getflavor()
    {
        return Engine.theEngine.session.mo_flavor;
    }

}