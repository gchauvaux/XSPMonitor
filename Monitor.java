package maxmatcher.monitor;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.sql.*;
import java.util.*;
import java.util.Date;

import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import javax.servlet.jsp.jstl.core.Config;

import org.apache.commons.lang.StringEscapeUtils;

//import com.sun.java.util.jar.pack.Package.Class.Method;



import javax.script.ScriptEngineFactory; 
import javax.script.ScriptEngineManager; 
import javax.script.ScriptEngine; 

import java.net.URLClassLoader;
import java.lang.ClassLoader;
import java.net.URL;

public class Monitor extends HttpServlet 
{
	public static boolean compare(Object o1,Object o2)
	{
		if(o1==null)
		{
			if(o2==null)
				return true;
			else
				return false;
		}
		return o1.equals(o2);
	}
	public static Integer asInteger(Object o)
	{
		if(o==null)
			return (Integer) null;
		try{
		return new Integer(Integer.parseInt((String)o));
		}catch(Exception e){return (Integer) null;}
	}
	public static char mycharAt(String s,int i)
	{
		try
		{
			return s.charAt(i);
		}
		catch(Exception e)
		{
			return 0;
		}
	}
	public void tlog(String peventtype,String pname,boolean pimmediate)
	{
		long lmilli;
		
		if(!hastolog(peventtype))
			return;
		if(mo_tlog==null)
			return;
		Date ldate=new Date();
		long ltime= System.currentTimeMillis( );
		String lmoment=ldate.toLocaleString() ;
		if(mo_time!=0)
		{
			lmilli=mo_time>0?ltime - mo_time:0;
			putLogData(lmoment,mo_user,lmilli,mo_eventtype,mo_name,"","",(String)Session.getAttribute("MO_SERVER"));
				
		}
		if(peventtype==null)  // fin chrono
			mo_time=0;
		else
		{
			mo_time=ltime;
			mo_eventtype=peventtype;
			mo_name=pname;
		}
		if(pimmediate)
		{
			putLogData(lmoment,mo_user,0,mo_eventtype,mo_name,"","",(String)Session.getAttribute("MO_SERVER"));
			mo_time=0;
		}
		
	}
	public void putLogData(String pmoment,String puser,long pmillisec,String pevent_type,String pname,String pretval,String pmessage,String pserver)
	{
		//Logger l=new Logger();
		//l.putLogData(pmoment, puser, pmillisec, pevent_type, pname, pretval, pmessage, pserver);
		try
	    {
	      Object[] parametre = {pmoment,puser,pmillisec,pevent_type,pname,pretval,pmessage,pserver};
	      Object resu=mo_tlogmethod.invoke(mo_tlog, parametre);
	      resu=resu;
	    }
	    catch (Exception e)
	    {
	    	String sss=e.getMessage();
	    	sss=sss;
	    }		
	}
	public void putLogData2(String pmoment,String puser,long pmillisec,String pevent_type,String pname,String pretval,String pmessage,String pserver)
	{
		   String JDBC_DRIVER = "com.mysql.jdbc.Driver";  
		   String DB_URL = "jdbc:mysql://localhost/MXMSEARCH";

		   //  Database credentials
		   String USER = "root";
		   String PASS = "";
		   
		   Connection conn = null;
		   PreparedStatement stmt = null;
		   try{
		      //STEP 2: Register JDBC driver
		      Class.forName("com.mysql.jdbc.Driver");

		      //STEP 3: Open a connection
		      System.out.println("Connecting to database...");
		      conn = DriverManager.getConnection(DB_URL,USER,PASS);

		      //STEP 4: Execute a query
		      System.out.println("Creating statement...");
		      stmt = conn.prepareStatement("insert into tlog(moment,userid,millisec,eventtype,name,retval,errmsg,server) values(NOW(),?,?,?,?,?,?,?)");
		      //stmt.setString(1, pmoment);
		      stmt.setString(1, puser);
		      stmt.setLong(2, pmillisec);
		      stmt.setString(3,  pevent_type);
		      stmt.setString(4, pname);
		      stmt.setString(5, pretval);
		      stmt.setString(6, pmessage);
		      stmt.setString(7, pserver);
		      stmt.execute();
		      stmt.close();
		      conn.close();
		   }catch(SQLException se){
		      //Handle errors for JDBC
		      se.printStackTrace();
		   }catch(Exception e){
		      //Handle errors for Class.forName
		      e.printStackTrace();
		   }finally{
		      //finally block used to close resources
		      try{
		         if(stmt!=null)
		            stmt.close();
		      }catch(SQLException se2){
		      }// nothing we can do
		      try{
		         if(conn!=null)
		            conn.close();
		      }catch(SQLException se){
		         se.printStackTrace();
		      }//end finally try
		   }//end try
		   System.out.println("Goodbye!");
	}	
	
	boolean hastolog(String peventtype)
	{
		if(peventtype==null)
			return true;
		if(tlevel=="0") return true;
		if(tlevel=="1") return false;
		if(tlevel=="2")  // debut/fin session
				if(peventtype!="D" && peventtype!="F")	return false;
				else return true;
		if(tlevel=="3")  // debut/fin session perf
				if(peventtype!="D" && peventtype!="F" && peventtype!="P") return false;
				else return true;
		return true;
	}
	void tlogtimer(String pcode,long ptime)
	{ 
		if(mo_tlog==null)
			return;
			
		String lmoment=(new Date()).toLocaleString() ;	
		putLogData(lmoment,mo_user,ptime,"M",pcode,"","",(String)Session.getAttribute("MO_SERVER"));
	}

	public void error(String message)
	{
		
	}
	public void doAbandonAppli()
	{
		Session.invalidate();
	}
	public void doEnd()
	{
		
	}
	public void save()
	{
		
	}
	public void restore()
	{
	
	}
	public void abandon()
	{
	
	}
	public class Endexception extends Exception
	{
		
	}
	public class Context
	{
		public String prg,name,caller;
		public int ptr,len,lastaff,lastpage;
		public Stack stack,laststack;
		public boolean issaved,request,transfer;
		
		
		Context(String pprg,String name,String pcaller)
		{
			prg=pprg;
			len=prg.length();
			ptr=0;
			stack=new Stack();
			laststack=new Stack();
			lastpage=-1;
			lastaff=-1;
			caller=pcaller;
			issaved=false;
			request=false;
			transfer=false;
		}
		public String getverb()
		{
			return getword2(true).toUpperCase(); // verbe pas donné par eval
		} 
		public String getword() // throws IOException,ServletException
		{
			return getword2(false);			
		}
		public String getword2(boolean noeval) // throws IOException,ServletException
		{
			//----------------------------------------------------------------------------
			//  recherche token. Si on trouve § et que noeval=null txt.. § on renvoit eval(txt..)
			//  noeval n'est positionné que par getverb
			  // ret,s,i,cont,inword,c,rpos;
			String ret,txt;
			int i,rpos;
			boolean cont,inword;
			char c;
			
			ret="";			// valeur retournée
  		  	i=ptr;
			rpos=i;			// sauve position départ
			cont=true;		// indic dans le token
			inword=false;
			while(i<len && cont)
			{
			    c=mycharAt(prg,i++);
			    switch (c)
			    {
			      case ' ':		// whitespaces
			      case '\t':
			      case '\n':
			      case '\r':
			        if (inword)
			          cont=false;  // fin si sortie de token
			        break;
			      case ';' :		// délimiteurs
			      case '(' :
			      case ')' :
			      case '=' :
			      case '*' :
			      case '§' :
			      case ',' :
			        if (inword)
			          i--;		// si dans token on ressort le délimiteur
			        else
			          ret=Character.toString(c);    // sinon le token est le délimiteur
			        cont=false; // dans le deux cas on a terminé
			        break;
			      default:		// autres cas on collecte
			        ret+=c;
			        inword=true;
			        break;
			    }
			  }
			  ptr=i;	// on positionne le pointeur d'instruction
			  if(!noeval && compare(ret,"§"))
			  {  // on évalue le javascript
			    ptr=rpos;		// retour à la position de départ
			    txt=geteval();   // lecture du javascript
			    try{
			      ret=eval(txt).toString();  // exécution
			    }  
			    catch(Exception e){}
			  }  
			 // ret=ret.toUpperCase();  // mise en majuscule
			  return ret;			
		}
		public String gettrans()
		{
			return getword().toUpperCase();
		}          
		
		public String geturl()
		{
		  String url,name;
		  int idx;
		  
		  url= getword();    // token
		  if(compare(url,";"))
		  { ptr--; 
		    return "";
		  }
		  if(url.indexOf(".")<0)
		    url+=".ASP";
		  mo_page=url;
		  if((name=pathfile2(url))==null)  // recherche dans MO_FILES
		    return url;             // non trouvé
		  return name;   // trouvé sans "?..." 
		}          
		public String pathfile2(String f)
		{ 
			return "/"+f;
		}

		//---------------------------------------------------------------------------- geteval
		public String geteval() // throws IOException,ServletException
		{
		//----------------------------------------------------------------------------
		// lecture § java ... §

		  int i1,i2;
		  
		  i1=prg.indexOf("§",ptr)+1;   // premier §
		  if((i2=prg.indexOf("§",i1))<0) // second 
			  try
		  {
			  ServerTransfer("mo_error.jsp");			// : pas trouvé
		  }
		  catch(Exception e){}
		  
		  ptr=i2+1;                 // on saute le second §
		  return prg.substring(i1,i2);
		 }
		public String skipcom()
		{
			//----------------------------------------------------------------------------
			// saut commentaires = jusqu'au cr/lf  inclus
			String s;
			int i,idep;
			char c,lf;

			s=prg;
			i=ptr;
			idep=i;
			lf=10;
			  
			while((c=mycharAt(s,i++))!=lf && c!=0);
			ptr=i;
			return s.substring(idep,i);
		}
		public void xpush(Ct pct)
		{
			stack.push(pct);
		}
		public Ct xpop()
		{
			Ct lct;
			try
			{
				lct=(Ct)stack.pop();
			}
			catch(Exception e)
			{
				lct=null;
			}
			return lct;
		}
		public String skip()
		{
			// saut de l'instruction = jusqu'au ";" non inclus en tenant compte des ()
			  String s;
			  int i,nbpar,ineval,idep;
			  boolean cont;
			  char c;

			  s=prg;
			  i=ptr;
			  idep=i;
			  cont=true;
			  nbpar=0;
			  
			  while(cont)
			  {
			    c=s.charAt(i++); 
			    switch(c){
			      case ';' :
			        if(nbpar==0 )
			          cont=false;
			        break;
			      case '(' :
			        nbpar++;
			        break;
			      case ')' :
			        nbpar--;
			        break;
			      case '§' :
			        ptr=i-1;
			        geteval();   // sauter §.....§
			        i=ptr;
			        break;
			      case '"' :
			        ptr=i-1;
			        getstring();   // sauter §.....§
			        i=ptr;
			        break;
			    }          
			  }
			  ptr=i-1;
			  return s.substring(idep,i);
		}
		public String getstring()
		{
			//  recherche chaîne "................" ou eval § ... §
			  String ret,s,del;
			  int i;
			  boolean cont,inword;
			  char c;
			  
			  ret="";			// valeur retournée
			  del="";
			  s=prg;
			  i=ptr;
			  cont=true;		// indic dans le token
			  while(cont)
			  {
			    c=s.charAt(i++);
			    switch (c)
			    {
			      case 0 :		// fin script
			        cont=false;
			        break;
			      case '"':
			      case '§':
			        if (compare(del,""+c))
			        {
			          cont=false;  // fin si sortie de chaine
			          break;
			        }
			        else
						if (compare(del,""))
						{
				          del=""+c;
				          break;
				        }
				        else
							if(compare(del,"\""))		// chaine "...............§ ".......§ "........" on a recu §
							{
								del="";
								break;
							}
			      default:		// autres cas on collecte
			        if(c=='\t')	// tab gène command (trim) mais est a conserver pour § comme separateur lexical
						c=' ';  // blanc devrait aller dans les 2 cas
			        if(!compare( del,""))
			          ret+=c;
			        break;
			    }
			  }
			  ptr=i;	// on positionne le pointeur d'instruction
			  if(compare(del,"§"))
			    try
			    {
			      ret=eval(ret).toString();  // exécution
			    }  
			    catch(Exception e){}
			  return ret;
		}
		public boolean option(String word){
			//----------------------------------------------------------------------------
			// lit un token optionnel si pas trouvé, restaure le pointeur et renvoie false
			  int i;
			  String w;
			  i=ptr;
			  w=getword().toUpperCase();
			  if(compare(word.toUpperCase(),w))
			    return true;
			  ptr=i;
			  return false;
			}
		public String getcond()
		{
			return getword();
		} 
		public Vector getlist()
		{
			//----------------------------------------------------------------------------
			//  recherche liste  (connexion si bds=true),.... facultative
			String src;
			Vector arr;
			int i;

			arr=new Vector();
			src=getword();
			if(compare(src,";"))
			{
				ptr--;   // on revient sur le ;
				 return arr;
			}
			while (true)
			{
			  if(!compare(src,",") && !compare(src,";") )
			  {
				arr.add(src);
				src=getword();
			  }
			  if(compare(src,";"))
			  {
			    ptr--;   // on revient sur le ;
			    return arr;
			  }
			  src=getword();
			}        
		}
			//---------------------------------------------------------------------------- getlist2
		public Vector getlist2()
		{
			//----------------------------------------------------------------------------
			//  recherche liste  pour CASE
			String src;
			Vector arr;
			int i=0;

			arr=new Vector();
			src=getword();		
			arr.add(src);
			i=ptr;
			while (true)
			{
				src=getword();		
				if(!compare(src,","))
				{
					ptr=i;   // on revient sur le ;
					return arr;
				}
				else
				{
					src=getword();		
					arr.add(src);
					i=ptr;
				}
			}        
			}

	}
	
	public enum VERB
	{
		REPEAT,BLOCK,CASE,CALL,ANY
	}
	public enum INSTR
	{
		ABANDON,RETURN,ALERT,TITLE,BP,TLOGCLASS,TLOGLEVEL,TLOG,TLOGBEG,TLOGEND,SQLTRACE,COMMANDS,NOCOMMANDS,LOADXML,SWAPXML,CLOSEXML,TRANSFORM,SENDDATA,DISP,
		SEND,EXEC,CALL,SAVE,RESTORE,TRANS,BRANCH,TRANSFER,CASE,CASE2,OTHER,REPEAT,CONTINUE,EVAL,EVAL2,READTRANS,DEFINEDATASOURCE,GETCONNECTION,
		CLOSECONNECTION,BEGINTRANS,COMMIT,ROLLBACK,ERROR,LOOP,BREAK,ENDMODAL,LANGUAGE,AGAIN,MODAL
	}
	public static Hashtable INSTRS; 
	public class Ct
	{
		public VERB verb;
		int data;
		Ct(VERB pverb,int pdata)
		{
			verb=pverb;
			data=pdata;
		}

	}
    public Object oeval=null;
	
	String stname,stinstr,mo_page;
	int stline;
	HttpSession Session;
	HttpServletRequest request;
	HttpServletResponse response;
	PrintWriter writer;
	ServletContext Application;
	Hashtable TRANS;
	int save_callstack,save_savestack;
	Object mo_tlog;
	java.lang.reflect.Method mo_tlogmethod;
	Boolean mo_islog;
	String tlevel;
	long mo_time;
	String mo_user,mo_eventtype,mo_name;
	String nasterlog="Y";
	
	public String Request(String pattribute)
	{
		return (String)request.getParameter(pattribute);
	}
	public Stack newArray()
	{
		return new Stack();
	}
	public Hashtable newarray2()
	{
		return new Hashtable();
	}
	public void ServerExecute (String ppage) throws IOException,ServletException
	{
		request.getRequestDispatcher(ppage).include(request,response);
	}
	public void MyExecute(String page)
	{
		try{
		ServerExecute(page);
		}catch(Exception e){error("myexecute("+page+")");}
	}

	public void ServerTransfer (String ppage) throws IOException,ServletException
	{
		request.getRequestDispatcher("/"+ppage).forward(request,response);
	}

	
	//-------------------------------------------------------------------------- READTRANS
	public void readtrans(String file)   // lecture d'un fichier de transactions
	//-------------------------------------------------------------------------- 
	{
		String txt="",name;
		int i;
		int c;
		Context temp;
		try
		{
			InputStream strm=Application.getResourceAsStream(file);
	
			while((c=strm.read())!=-1)
				txt+=(char)c;
			strm.close();
		}
		catch(Exception e)
		{
			String s=e.getMessage();
		}
		
		temp=new Context(txt,"","");

		txt=txt.toUpperCase();   // en majuscules
		while((i=txt.indexOf("DEFTRANS",temp.ptr))>=0)  // début définition script
		{
			  temp.ptr=i+8;
			  name=temp.getword().toUpperCase();     // nom transaction
			  temp.getword().toUpperCase();          // ";"
			  i=txt.indexOf("ENDDEF",temp.ptr);   // recherche fin
			  if(i<0)
					break;   // pas de fin : non traité
			  TRANS.put(name,temp.prg.substring(temp.ptr,i));
			  i+=6;
			  temp.ptr=i;		
		}
	} 
	
	public String texttrans(String t)
	{
		//var p=TRANS(t);  
		String p=(String)TRANS.get(t);  // dicodict
	    return p;
	}
	
	public void initINSTRS()
	{
		INSTRS=new Hashtable();
        INSTRS.put("ABANDON",			INSTR.ABANDON);
		INSTRS.put("RETURN",			INSTR.RETURN);
		INSTRS.put("ALERT",		    	INSTR.ALERT);
		INSTRS.put("TITLE",			    INSTR.TITLE);
		INSTRS.put("BP",	    		INSTR.BP);
		INSTRS.put("TLOGCLASS",	    	INSTR.TLOGCLASS);
		INSTRS.put("TLOGLEVEL",	    	INSTR.TLOGLEVEL);
		INSTRS.put("TLOG",		    	INSTR.TLOG);
		INSTRS.put("TLOGBEG",			INSTR.TLOGBEG);
		INSTRS.put("TLOGEND",			INSTR.TLOGEND);
		INSTRS.put("SQLTRACE",			INSTR.SQLTRACE);
		INSTRS.put("COMMANDS",			INSTR.COMMANDS);
		INSTRS.put("NOCOMMANDS",		INSTR.NOCOMMANDS);
		INSTRS.put("LOADXML",			INSTR.LOADXML);
		INSTRS.put("SWAPXML",			INSTR.SWAPXML);
		INSTRS.put("CLOSEXML",			INSTR.CLOSEXML);
		INSTRS.put("TRANSFORM",			INSTR.TRANSFORM);
		INSTRS.put("SENDDATA",			INSTR.SENDDATA);
		INSTRS.put("DISP",		    	INSTR.DISP);
		INSTRS.put("SEND",		    	INSTR.SEND);
		INSTRS.put("EXEC",		    	INSTR.EXEC);
		INSTRS.put("CALL",		    	INSTR.CALL);
		INSTRS.put("SAVE",		    	INSTR.SAVE);
		INSTRS.put("RESTORE",	    	INSTR.RESTORE);
		INSTRS.put("TRANS",		    	INSTR.TRANS);
		INSTRS.put("BRANCH",			INSTR.BRANCH);
		INSTRS.put("TRANSFER",			INSTR.TRANSFER);
		INSTRS.put("CASE",			    INSTR.CASE);
		INSTRS.put("CASE2",	 	    	INSTR.CASE2);
		INSTRS.put("OTHER",		    	INSTR.OTHER);
		INSTRS.put("REPEAT",			INSTR.REPEAT);
		INSTRS.put("CONTINUE",			INSTR.CONTINUE);
		INSTRS.put("EVAL",			    INSTR.EVAL);
		INSTRS.put("EVAL2",		    	INSTR.EVAL2);
		INSTRS.put("READTRANS",	        INSTR.READTRANS);
		INSTRS.put("DEFINEDATASOURCE",	INSTR.DEFINEDATASOURCE);
		INSTRS.put("GETCONNECTION",	    INSTR.GETCONNECTION);
		INSTRS.put("CLOSECONNECTION",	INSTR.CLOSECONNECTION);
		INSTRS.put("BEGINTRANS",		INSTR.BEGINTRANS);
		INSTRS.put("COMMIT",		    INSTR.COMMIT);
		INSTRS.put("ROLLBACK",		    INSTR.ROLLBACK);
		INSTRS.put("ERROR",	        	INSTR.ERROR);
		INSTRS.put("LOOP",	        	INSTR.LOOP);
		INSTRS.put("BREAK",		        INSTR.BREAK);
		INSTRS.put("ENDMODAL",		    INSTR.ENDMODAL);
		INSTRS.put("LANGUAGE",		    INSTR.LANGUAGE);
		INSTRS.put("AGAIN",		    	INSTR.AGAIN);
		INSTRS.put("MODAL",		    	INSTR.MODAL);
		
	}
	public Object eval(String s)
	{
		ScriptEngineManager manager = new ScriptEngineManager(); 
		ScriptEngine engine = manager.getEngineByName("JavaScript");
	    engine.put("Session",Session);
	    engine.put("request",request);
	    try{
		    oeval=engine.eval(s);
	    }
	    catch(Exception e)
	    {
	    	String ss=e.getMessage();
	    	ss=ss;
	    }
	    return oeval;   
	}

	public void doGet(HttpServletRequest preq, HttpServletResponse presp)
			throws ServletException, IOException 
	{
		Integer npagein,npageout;
		String r;
		Context ctx;

		//ClassLoader originalClassLoader = Thread.currentThread().getContextClassLoader();
		//ClassLoader newClassLoader = new JarSeekingURLClassLoader(new File(""));
 
/*
		try {    
		    //Thread.currentThread().setContextClassLoader(newClassLoader);
		    // write code to load new classes
		    Class.forName("maxmatcher.moniteur.JarSeekingURLClassLoader"); //JarSeekingURLClassLoader
		}
		catch(Exception e){
			String s=e.getMessage();
			s=s;
		}
		finally 
		{
		    Thread.currentThread().setContextClassLoader(originalClassLoader);
		}
	*/	
		try
		{
			mo_page="";
			request=preq;
			response=presp;
			writer =presp.getWriter();
			Application=getServletContext();
			Object o=Application.getAttribute("TRANS");
			if(o==null)
			{
				o=new Hashtable();
				Application.setAttribute("TRANS", o);
				TRANS=(Hashtable) o;	
				readtrans("/moniteur.mon");
			}
			else
				TRANS=(Hashtable) o;		
			
			initINSTRS();
			
			String tmp;
			int i;
			
			boolean bfirst=preq.getParameter("mode")!=null;
			
			//Hashtable TRANS;						 // dico transactions
			Session=preq.getSession();
			String sid=Session.getId();
			Session.setAttribute("MO_ERRMESS", "");
			Session.setAttribute("mo_inerror",null);
			Session.setAttribute("MO_ALERT", 0);
			Session.setAttribute("MO_NOMPAGE","");
			String mo_rc="";
			String site=new File(".").getCanonicalPath();
	
			//-----------------------------------------------------------------------------
			//traitement des erreurs : cas ou moniteur.asp est la page de traitement 500-100
			//-----------------------------------------------------------------------------
	/*
			err=Server.GetLastError();
			if(err.Number!=0)
			  error("1 GetLastError");
			
	*/
			boolean bForm=(compare(preq.getMethod(),"POST"));
			String req=(String)Session.getAttribute("MO_REQUEST");	// en principe, le moniteur n'est appelé que d'une
			Session.setAttribute("MO_REQUEST","");   // page affichée et la requête est dans Request.Form
	
			if(   Session.getAttribute("MO_SAVESTACK")==null
					   || bfirst
					   || compare(req.toUpperCase(),"INIT")
					   || compare(Request("init"),"yes"))
			{
				  Session.setAttribute("MO_DOCTYPE","<HEAD><meta http-equiv=\"expires\" content=\"-1\"></HEAD>");
				  Session.setAttribute("MO_SAVESTACK",newArray());
				  Session.setAttribute("MO_CALLSTACK",newArray());
				  Session.setAttribute("MO_RETURN","");
				  Session.setAttribute("MO_ACTION",""); 
				  Session.setAttribute("MO_RESULT",null);
				  Session.setAttribute("MO_COMMANDES","");
				  Session.setAttribute("MO_INTRANS",false); 
				  Session.setAttribute("MO_LANGUE","undefined"); //"fr"; //"en";
				  Session.setAttribute("MO_NOSAVE",false); // sauvegarde forme
				  Session.setAttribute("MO_CLEAR",false); // clear objet xmlservice
				  Session.setAttribute("MO_LASTRETURN","");  // dernier code retour
				  Session.setAttribute("MO_APPLICATION","NASTER");
				  Session.setAttribute("MO_TRANSACTION","NASTER");
				  Session.setAttribute("MO_CALLER","");
				  Session.setAttribute("MO_USER","");
				  Session.setAttribute("EX_UTILISATEUR","");
				  Session.setAttribute("EX_UTILISATEUR2","");
				  Session.setAttribute("MO_HASERROR",false);
				  Session.setAttribute("MO_NOALERTERROR",false);
				  Session.setAttribute("MO_WRNAME","PLWRAPPER");
				  Session.setAttribute("MO_MENU","NASTER");
				  Session.setAttribute("MO_OLD_MENU","");
				  Session.setAttribute("MEM_GL_BLINKTIME",0);
				  Session.setAttribute("MEM_GL_WIDTH",0);
				  Session.setAttribute("MO_NEXTDATE",null);
				  Session.setAttribute("MO_NEXTCOMMENT","");
				  Session.setAttribute("MO_ISLOG",false);
				  tmp								=request.getServerName();
				  i									=tmp.indexOf(".");
				  if(i>=0)
					tmp								=tmp.substring(0,i);
				  //tmp								=tmp.substring(0,10);
				  Session.setAttribute("MO_SERVER",tmp+".mcsfr.net");		//+++++++++++++++++++
				  Session.setAttribute("MO_VBSERR",false);
				  Session.setAttribute("MO_VIAXML",0);      // 0=sans 1=transform 2=send avec xml
				  Session.setAttribute("MO_SENT_TRANS","");
				  Session.setAttribute("MO_INSTANCE","");
				  Session.setAttribute("MO_SCHEMA","");
				  Session.setAttribute("MO_PASSWORD","");
				  Session.setAttribute("MO_APPLI","");
				  Session.setAttribute("MO_MESSLU",false);
				  Session.setAttribute("MO_MESSEXT",true);
				  Session.setAttribute("MO_STARTMENU","TRANS MENU");
				  Session.setAttribute("MO_NOINCR",false);
				  Session.setAttribute("MO_TLOG",null);
				  Session.setAttribute("MO_TLOGMETHOD",null);
				  
				  tmp=Request("instance");
				  if(tmp!=null)
					Session.setAttribute("MO_INSTANCE",tmp); 
				  if(compare(tmp,"recec"))
					Session.setAttribute("CO_STYLE","STBASE");
				  else
					Session.setAttribute("CO_STYLE","ST0001");
				  tmp=Request("schema");
				  if(tmp!=null)
					Session.setAttribute("MO_SCHEMA",tmp); 
					
				  tmp=Request("password");
				  if(tmp!=null)
					Session.setAttribute("MO_PASSWORD",tmp); 
					
				  tmp=Request("instance2");	// instance Oracle - serveur d'impression
				  if(tmp!=null)
					Session.setAttribute("MO_INSTANCE2",tmp); 
				  
				  tmp=Request("schema2");
				  if(tmp!=null)
					Session.setAttribute("MO_SCHEMA2",tmp); 
					
				  tmp=Request("password2");
				  if(tmp!=null)
					Session.setAttribute("MO_PASSWORD2",tmp); 
					
				  tmp=Request("appli");
				  if(tmp!=null)			//  nom appli pour ETATS Crystal TSTNSD migre TSTNS3 prod 
					Session.setAttribute("MO_APPLI",tmp);
				  else
					Session.setAttribute("MO_APPLI","TSTNS3"); //"TSTNSD";
					
				  tmp=Request("wrname");
				  if(tmp!=null)			//  nom wrapper
					Session.setAttribute("MO_WRNAME",tmp);
					
				  tmp=Request("mode");
				  Session.setAttribute("MO_MODE",tmp); 
				  tmp=Request("flavor");
				  if(tmp!=null)			// appli = CTX	
					Session.setAttribute("MO_FLAVOR",tmp); 
				  else
					Session.setAttribute("MO_FLAVOR",""); 
		
				  tmp=Request("servimp");
				  if(tmp!=null)			// serveur impression 
					Session.setAttribute("MO_SERVIMP",tmp); 
				  else
					Session.setAttribute("MO_SERVIMP","mcs20c"); 
		
				  tmp=Request("droit");
				  if(tmp!=null)			// ILLIMITE cnil
					Session.setAttribute("MO_DROIT",tmp); 
				  else
					Session.setAttribute("MO_DROIT",""); 
				      
				  String tmptrans=Request("trans");
				  if(tmptrans!=null)	
					Session.setAttribute("MO_STARTMENU","TRANS "+tmptrans+";");
				      
				  tmp=Request("startmenu");
				  if(tmp!=null)			// appli = CTX	
					Session.setAttribute("MO_STARTMENU",tmp); 
		
				  tmp=Request("noincr");
				  if(tmp!=null)			// appli = CTX	
					Session.setAttribute("MO_NOINCR",true); 
		
				  if(Session.getAttribute("MO_NPAGEOUT")==null)
				    Session.setAttribute("MO_NPAGEOUT",0);
				  if(req==null)
				  {
				    Session.setAttribute("MO_REQUEST","");
				    req="";
				  }
				  //TRANS= MO_TRANS;
				  Session.setAttribute("MO_SESSTRANS",false);    
		
				  req=Request("mo_request");
				  String tr;
				  if(req==null)  
				  {
					if(tmptrans==null)
					   tr="DEFAULT";
					else
					   tr=tmptrans;
					req="TRANS "+tr+";";
				  }
				  
				  Session.setAttribute("MO_TIME",0);		                           
				  Session.setAttribute("MO_MODAL","NO");
			}				
			else
		  
			//------------------------------------------------------------------------------
			//  appel par SUBMIT (POST ou GET) à partir d'une page affichée
			//------------------------------------------------------------------------------
			{
				  mo_tlog=Session.getAttribute("MO_TLOG");
				  mo_tlogmethod=(java.lang.reflect.Method)Session.getAttribute("MO_TLOGMETHOD");
				  tlevel=(String)Application.getAttribute("MO_TLOGLEVEL");
				  mo_islog=(Boolean)Session.getAttribute("MO_ISLOG");

				//------------------------ CONTROLE DE LA NAVIGATION -------------------------
			
				  if(preq.getParameter("noincr")==null && !(Boolean)Session.getAttribute("MO_NOINCR"))  //  incrémenter                      //kkk
				  {
					    npagein= asInteger(preq.getParameter("MO_NPAGE"));  // numéro de page généré dans le HTML
					    npageout=(Integer)Session.getAttribute("MO_NPAGEOUT");
					   // npageout= Session.getAttribute("MO_NPAGEOUT");        // numéro en cours
					    if(npageout==null)
							npageout=npagein;
							
					    if (npagein!=null && npagein!= npageout)
					    {			     // si # on ne renvoie rien
						      presp.setStatus(204);
						      Session.setAttribute("F_InControl", false);
							  Session.setAttribute("F_InPage", false);
						      return;
					    } 
					    
					    // incrémentation du numéro de page
					    Session.setAttribute("MO_NPAGEOUT",npageout+1);
				  }
				  r=(String)preq.getParameter("MO_RESULT");
				  if(r!=null)							// éviter modifier MO_RESULT si pas dans la forme car non sauvé dans un CALL
					Session.setAttribute("MO_RESULT",r);
				  r=(String)preq.getParameter("MO_REQUEST");					// requête
				  if(r!=null && !compare(r,"") )
				    req=r;
				  r=(String)preq.getParameter("MO_RETURN");
				  if(r!=null && !compare(r,""))
				  {
				    Session.setAttribute("MO_RETURN", r);  // code retour
				    Session.setAttribute("MO_ACTION", r);  
				  }
				  else
				    Session.setAttribute("MO_ACTION", "");  
				  
			
				  Session.setAttribute("MO_CLEAR",false);
			}
		
		//------------------------------------------------------------------------------
	//	    TRAITEMENT DES SCRIPTS
		//------------------------------------------------------------------------------
		// Session(MO_CTX)= contexte du script en cours
		
			while(true)														// boucle sur retour au menu départ
			{
				save_callstack=((Stack)Session.getAttribute("MO_CALLSTACK")).size();
				save_savestack=((Stack)Session.getAttribute("MO_SAVESTACK")).size();
				if(req !=null && !compare(req,""))
				{    // si il y a une requête (req)
					 // pour une fenetre modale on sauve le niveau des piles en cas de sortie anormale (fermeure par bouton systeme)
				// sauvegarde script en cours si nécessaire
				  if(Session.getAttribute("MO_CTX")!=null)
				  {
				    ctx=(Context)Session.getAttribute("MO_CTX");
				    ((Stack)Session.getAttribute("MO_CALLSTACK")).push(ctx); // push(arr,x) remplace arr.push(x) (bug)
				  }
				// initialisation script
				  ctx=new Context(req,"request","");
				  ctx.request=true;
				  Session.setAttribute("MO_CTX",ctx);
				}   
		
				//boucle sur les scripts imbriqués
		
				mo_eventtype="";
				mo_name="";
				mo_time=0;
				mo_user=(String)Session.getAttribute("EX_UTILISATEUR");
				mo_page="";
				
			//try
			//{
				while(true)				// boucle sur transactions empilées
				{
				  Context lctx=(Context)Session.getAttribute("MO_CTX");
				  if(lctx!=null)
				  {
					  play(lctx);   // exécution du script
					  tlog("T","/"+((Context)Session.getAttribute("MO_CTX")).name+"/",true);
				  }
				  if(((Stack)Session.getAttribute("MO_CALLSTACK")).size()==0)
				    break;
				  if(((Context)Session.getAttribute("MO_CTX")).request && ((Context)Session.getAttribute("MO_CTX")).transfer)
				  {
					  if(((Stack)Session.getAttribute("MO_CALLSTACK")).size()==0)
						break;
					  ((Stack)Session.getAttribute("MO_CALLSTACK")).pop();
				  }
				  Session.setAttribute("MO_CTX",((Stack)Session.getAttribute("MO_CALLSTACK")).pop()); // script suivant dans la pile
				  if(((Context)Session.getAttribute("MO_CTX")).issaved)
				  {
					restore();
			        ((Context)Session.getAttribute("MO_CTX")).issaved=false;
			      }
				  Session.setAttribute("MO_TRANSACTION",((Context)Session.getAttribute("MO_CTX")).name);
				  Session.setAttribute("MO_CALLER",((Context)Session.getAttribute("MO_CTX")).caller);
				}
				Session.setAttribute("MO_CTX",null);
				req=(String)Session.getAttribute("MO_STARTMENU");	
			    Session.setAttribute("MO_RETURN","");
			    break;
			}	
		}
		catch(Endexception e)
		{
			
		} 
		writer.flush();
		writer.close();
	}
	
	
	
	
	public void play(Context mo_ctx) throws Endexception
	{
		  if(mo_ctx==null)
			  return;
		  String mo_rc="",lab,tmp,nptr,verb,k,s,p,t,txt,m,r;
		  String x,file,a,cnx,cmd,src,ds,a2,wrapp,ors,w,oxmls,name,commandes,lang,langa;
		  String ltr,cnt,plang,wrp,C="",page;
		  Context nctx;
		  char c;
		  int i,mpos;
		  boolean ll,mo_skipcase,mo_more,bfound,isin,bcon2,cont,not;
		  Ct ct;
		  INSTR iverb;
		  Hashtable d;
		  Vector arr;
		  

		  Session.setAttribute("MO_TRANSACTION",mo_ctx.name);
		  Session.setAttribute("MO_CALLER",mo_ctx.caller);
	
		  mo_skipcase=false;   // indicateur : sauter les ordres CASE après un CASE positif
		  mo_more=true;        // indicateur : continuer la boucle

		  while(mo_more)
		  {
		    
		    //------------------------TRAITEMENT CODES RETOUR-------------------------
		    
		    tmp=((String)Session.getAttribute("MO_RETURN")).toUpperCase( ); // code retour
		    if(tmp!=null && !compare(tmp,""))
		    {	// si le code retour est à blanc, mo_rc n'est pas modifié
		      if (tmp==null || !compare(tmp,"ERROR"))
				Session.setAttribute("MO_LASTRETURN",tmp);    
		      i=tmp.length();
		      while(--i>0)		// élimination blancs et ;
		        if((c=tmp.charAt(i))!=' ' && c!=';')
		          break;
		      tmp=tmp.substring(0,i+1);
		      mo_rc=tmp;

		        if (compare(mo_rc, "ERROR"))		// erreur : retour vers la dernière page affichée
		        {
		            if(mo_ctx.lastaff==-1)
		               return;		// retour script supérieur
		            mo_ctx.ptr=mo_ctx.lastaff;
		   //mo_ctx.stack est la pile d'exécution
		   //mo_ctx.laststack est la sauvegarde correspondante à la dernière page affichée
		            mo_ctx.stack=(Stack)mo_ctx.laststack.clone();
		        }
		        else if(compare(mo_rc,"AGAIN"))
		        {
		            // réexécution de la dernière page exécutée
		            if(mo_ctx.lastpage<0) // retour script supérieur
		              return;
		            mo_ctx.ptr=mo_ctx.lastpage;
		        }          
		        else if(compare(mo_rc,"LOOP") || compare(mo_rc,"BREAK"))
		        {
		        	// LOOP retour en début de boucle REPEAT
		            // BREAK fin de boucle
		                // on recherche dans la pile d'exécution un ordre REPEAT
		            bfound=false;
		            while((ct=mo_ctx.xpop())!=null )
		            {
		              if(ct.verb==VERB.REPEAT)
		              {
		                mo_ctx.ptr=ct.data;  // on se positionne en début de boucle
		                if(compare(mo_rc,"BREAK"))   // si BREAK
		                  mo_ctx.skip();    // on saute l'ordre REPEAT 
		                bfound=true;
		                break;
		              }  
		            }
		            if(!bfound)
		              return;			// non trouvé on sort
		        }
		        else if(compare(mo_rc,"ERRSYS"))
		        {
		            error("6 erreur systeme");
		        }
		        else if(compare(mo_rc,"RETURN"))
		        {
		            mo_more=false;
		            if(!compare(mo_ctx.name,"request"))
		            {
		              Session.setAttribute("MO_RETURN","");  
		              mo_rc="";
		            }
		            continue;
		        }
		        else if(compare(mo_rc,"ABANDON"))
		        {
		          // arrêt session
					doAbandonAppli();
					return;
		        }
		        else if(compare(mo_rc,"ABANDON"))
		        {
				   // fin fenêtre modale
				    if(!compare(Session.getAttribute("MO_MODAL"),"NO"))
				    {
						if(((Stack)Session.getAttribute("MO_CALLSTACK")).size()>0)
						{
							Session.setAttribute("MO_CTX",((Stack)Session.getAttribute("MO_CALLSTACK")).pop()); 
							if(((Context)Session.getAttribute("MO_CTX")).issaved)
							{
								restore();
								((Context)Session.getAttribute("MO_CTX")).issaved=false;
							}
							Session.setAttribute("MO_TRANSACTION",((Context)Session.getAttribute("MO_CTX")).name);
							Session.setAttribute("MO_CALLER",((Context)Session.getAttribute("MO_CTX")).caller);
						}		    
						Session.setAttribute("MO_MODAL","NO");
/*
 			<script>
 				document.cookie = "modal=NO;path="+escape("/");
				document.cookie = "result=<%=escape(Session.getAttribute('MO_RESULT'))%>;path="+escape("/");
				window.close();
			</script>
*/
						Session.setAttribute("MO_NPAGEOUT",null);
						doEnd();  
					}
		        }
				else
				{
					c=mo_rc.charAt(0);
					if(c=='>')			//	><nom transaction>   (TRANSFER)
					{
						t=mo_rc.substring(1);
						Session.setAttribute("MO_MENU",t);												
						tlog("T",t,true);  //
						p=texttrans(t);  // texte du script
						if(p==null)
							error("7 Transaction "+t+" inexistante");
						nctx=new Context(p,t,"");   // création du contexte d'exécution   
						nctx.transfer=true;
						nctx.request=mo_ctx.request;
						Session.setAttribute("MO_CTX",nctx);  // le nouveau contexte est conservé dans la session				
						Session.setAttribute("MO_RETURN",""); 
						play(nctx);  // on l'exécute en récursif .   Au cas où l'on revient : 
						if(nctx.request)
							((Stack)Session.getAttribute("MO_CALLSTACK")).pop();
						if(((Stack)Session.getAttribute("MO_CALLSTACK")).size()>0)
							mo_ctx=(Context)((Stack)Session.getAttribute("MO_CALLSTACK")).pop(); // on dépile l'ancien contexte
						if(mo_ctx.issaved)
						{
							restore();
							mo_ctx.issaved=false;
						}       
						Session.setAttribute("MO_TRANSACTION",mo_ctx.name);
						Session.setAttribute("MO_CALLER",mo_ctx.caller);
						Session.setAttribute("MO_CTX",mo_ctx);    // on le stocke dans la session
					}		    
					else if(c=='@')			//	@<nom transaction>   (BRANCH)
					{
						t=mo_rc.substring(1);
						Session.setAttribute("MO_MENU",t);												
						tlog("T",t,true);
						p=texttrans(t);  // texte du script
						if(p==null)
							error("8 Transaction "+t+" inexistante");
						nctx=new Context(p,t,"");
						Session.setAttribute("MO_CTX",nctx);  // le nouveau contexte est conservé dans la session				
						Session.setAttribute("MO_SAVESTACK",newArray());//Session.setAttribute("MO_SAVESTACK").name="MO_SAVESTACK";
						Session.setAttribute("MO_CALLSTACK",newArray());//Session("MO_CALLSTACK").name="MO_CALLSTACK";
						Session.setAttribute("MO_RETURN",""); 
						play(nctx);  // on l'exécute en récursif .   Au cas où l'on revient : 
						if(nctx.request)
							((Stack)Session.getAttribute("MO_CALLSTACK")).pop();
						if(((Stack)Session.getAttribute("MO_CALLSTACK")).size()>0)
							mo_ctx=(Context)((Stack)Session.getAttribute("MO_CALLSTACK")).pop(); // on dépile l'ancien contexte
						if(mo_ctx.issaved)
						{
							restore();
							mo_ctx.issaved=false;
						}       
						Session.setAttribute("MO_TRANSACTION",mo_ctx.name);
						Session.setAttribute("MO_CALLER",mo_ctx.caller);
						Session.setAttribute("MO_CTX",mo_ctx);    // on le stocke dans la session
					}		    
					else if(c=='=')			//	=<nom transaction>   (CALL)
					{
						t=mo_rc.substring(1);
						Session.setAttribute("MO_OLD_MENU",Session.getAttribute("MO_MENU"));												
						Session.setAttribute("MO_MENU",t);												
						tlog("T",t,true);

		// les instructions marquées CALL servent a appeler t par "CALL TRANS t" au lieu de "BRANCH"

						mo_ctx.issaved=true;  // CALL
						mo_ctx.xpush(new Ct(VERB.CALL,mo_ctx.lastpage));
				        Session.setAttribute("MO_LASTRETURN","");    // CALL
				        Session.setAttribute("REINIT",1);				// testé au retour pour pgl qui lit directement Request.Form
						save();		// CALL 


						p=texttrans(t);  // texte du script
						if(p==null)
							error("9 Transaction "+t+" inexistante");
						nctx=new Context(p,t,(String)Session.getAttribute("MO_TRANSACTION"));
						Session.setAttribute("MO_CTX",nctx);  // le nouveau contexte est conservé dans la session				
						
						((Stack)Session.getAttribute("MO_CALLSTACK")).push(mo_ctx);  // CALL

						Session.setAttribute("MO_RETURN",""); 
						play(nctx);  // on l'exécute en récursif .   Au cas où l'on revient : 
						if(((Stack)Session.getAttribute("MO_CALLSTACK")).size()>0)
							mo_ctx=(Context)((Stack)Session.getAttribute("MO_CALLSTACK")).pop(); // on dépile l'ancien contexte
						if(mo_ctx.issaved)
						{
							restore();
							mo_ctx.issaved=false;
						}       
						Session.setAttribute("MO_TRANSACTION",mo_ctx.name);
						Session.setAttribute("MO_CALLER",mo_ctx.caller);
						Session.setAttribute("MO_CTX",mo_ctx);    // on le stocke dans la session
					}		    
		       }
		    
		      
		    }
		    Session.setAttribute("MO_RETURN","");  // raz code retour. mo_rc est conservé
		    
		    mpos=mo_ctx.ptr;		// position avant l'instruction 
		    
		    if(compare((verb =mo_ctx.getverb()),"")) // lecture du token suivant
		      break;						// si rien : fin de script

		    if(compare(verb,"*"))					// commentaire +++
		    {
		      mo_ctx.skipcom();
		      continue;
		    }
			   
		    if(compare(verb,";"))
		    {                 // FIN INSTRUCTION
		        cont=true;				// on examine la pile 
		        while(cont && (ct=mo_ctx.xpop())!=null )
		        {
		          switch(ct.verb)
		          {
		            case BLOCK :			//  parenthèse ouvrante : on est sorti de
		              mo_ctx.xpush(new Ct(VERB.BLOCK,0)); // l'instruction en cours 
		              cont=false;         // on réempile et on termine la boucle
		              break;  
		            case CASE :  // fin instruction CASE : les CASE suivants sont sautés
		              mo_skipcase=true;
		              break;
		            case REPEAT :  //retour en début de boucle
		              mo_ctx.ptr=ct.data;
		              mo_skipcase=false;  // tout ordre non CASE rétablit les CASE suivants
		              cont=false;
		              break;
		            case CALL :
		              mo_skipcase=false;
		              //restore();  
					  //mo_ctx.issaved=false;
		              break;
		          }    
		        } 
		        continue;               
		    }      

		    if(compare(verb,")"))
		    {              // dépilement de la parenthèse :"("
		      mo_ctx.xpop();
		      continue;
		    }
		        
		    if(compare(verb,"("))
		    {               // empilement du début de bloc
		      mo_skipcase=false;
		      mo_ctx.xpush(new Ct(VERB.BLOCK,0));
		      continue;
		    }
		    
		    bcon2=false;
		    
		    iverb=(INSTR)INSTRS.get(verb); //
		    switch (iverb) 
		    {
				//---------------------------TLOGCLASS "classe"  $$$$-----------------------------------
			  case TLOGCLASS:
				C=mo_ctx.getstring();
				try
			    {
			      Class clog=Class.forName (C);
			      mo_tlog=clog.newInstance ();
//					mo_tlog.putLogData(string lmoment,string mo_user,long lmilli,string mo_eventtype,string mo_name,string "",string "",string Session.getAttribute("MO_SERVER"));
			      Class [] typeParametre = {String.class,String.class,long.class,String.class,String.class,String.class,String.class,String.class};
			      mo_tlogmethod = clog.getMethod("putLogData",typeParametre);
   				  Session.setAttribute("MO_TLOG",mo_tlog);
   				  Session.setAttribute("MO_TLOGMETHOD",mo_tlogmethod);
			    }
			    catch (Exception e)
			    {
			      // La classe n'existe pas
					error("TLOGCLASS "+e.getMessage());
			    }
				
				break;

				//---------------------------TLOG ON/OFF/START/END log des temps passés dans la table tlog-----------
			  case TLOG:   
				C=mo_ctx.getword().toUpperCase();
				if(compare(nasterlog,"N"))
					break;								// pas de LOG !!!!!!!!!!!!!!
				if(compare(C,"ON"))		// démarrage
				{
						Session.setAttribute("MO_ISLOG",true);
						mo_islog=true;
						tlog("O","TLOG ON******",true);	
				}
				else if(compare(C,"START"))		// log user
				{
					if(mo_islog)
					{
						tlog("D","DEBUT SESSION ******",true);
					}
				}
				else if(compare(C,"END"))		// fin session
				{
					if(mo_islog)
					{
						tlog("F","FIN SESSION ******",true);	
					}
				}
				else			// arret OFF
				{
					if(mo_islog)
					{
						tlog("O","/TLOG OFF******/",true);	
						Session.setAttribute("MO_ISLOG",false);
						mo_islog=false;
					}
				}
				break;
		//---------------------------TLOGBEG nom : timing-----------
			  case TLOGBEG:   
				C=mo_ctx.getword();
				Session.setAttribute(C,(new Date()).getTime());
				break;
				
		//---------------------------TLOGBEG nom : timing-----------
			  case TLOGEND:   
				C=mo_ctx.getword();
				Object o=Session.getAttribute(C);
				if(o!=null)
				{
					Long tt=(Long)o;
					tlogtimer(C,(new Date()).getTime()-tt);
					Session.setAttribute(C,null);
				}
				break;

				
				
		      case ABANDON :  // arrêt session
					doAbandonAppli();
		     case RETURN :
		        mo_more=false;
		        if(!compare(mo_ctx.name,"request"))
		          Session.setAttribute("MO_RETURN","");   
		        continue;
		//---------------------------ALERT "message"-----------------------------------
			  case ALERT:
				C=mo_ctx.getstring();
				writer.write("<script>alert(unescape('"+StringEscapeUtils.escapeHtml(C)+"'));</script>"); 
				break;
		//---------------------------TITLE "titre"-----------------------------------
			  case TITLE:
				C=mo_ctx.getstring();
			    lang=mo_ctx.getword().toUpperCase();
			    langa=((String)Session.getAttribute("MO_LANGUE")).toUpperCase();
			    if(compare(lang,";"))
					mo_ctx.ptr--;
				if(compare(lang,";") && compare(langa,"UNDEFINED") || compare(langa,lang))
				{
					writer.write("<html><head><title>NASTER V1.0 "+C+"</title></head>");
					Session.setAttribute("MO_TITLE",C);
				}
				break;
		//---------------------------DEBUG-------------------------------------------
			  case BP:
				break;
		//---------------------------------COMMANDES "libellé=option/..." par langue ---------------COMMANDES
			  case COMMANDS:
			    commandes=mo_ctx.getstring();
			    lang=mo_ctx.getword().toUpperCase();
			    if(compare(lang,";"))
			    {
					Session.setAttribute("MO_COMMANDES",commandes);
					//lang="UNDEFINED";
					mo_ctx.ptr--;
				}
				else if(compare(((String)Session.getAttribute("MO_LANGUE")).toUpperCase(),lang))
					Session.setAttribute("MO_COMMANDES",commandes);
			    break;
		//---------------------------------NOCOMMANDES  --------------------------------------------NOCOMMANDES
			  case NOCOMMANDS:
				Session.setAttribute("MO_COMMANDES",null);
			    break;
		//---------------------------------DISP page----------------------------------------- DISP
		      case DISP :		// exécution d'une page d'affichage
		               // la page doit être suivie d'autres pages et du SEND
		               // ex : disp menu; exec aff; exec commandes; send; 
		               // sinon elle peut renseigner le code de requête ou de retour
		//----------------------------------SEND [nosave][clear][noreturn][nocontent]--------------------------------------- SEND        
		      case SEND :		 // envoi de la page
				Session.setAttribute("MO_NOSAVE",mo_ctx.option("nosave")); // pas de sauvegarde forme
				Session.setAttribute("MO_CLEAR",mo_ctx.option("clear")); // clear xmls
		        Session.setAttribute("F_InControl", false);
		        Session.setAttribute("F_InPage", false);
				// transaction dans une fenetre showmodaldialog : on ne revient pas
				// l'option NORETURN permet de rétablir les stacks
				if(mo_ctx.option("noreturn"))
				{
					if(save_callstack<((Stack)Session.getAttribute("MO_CALLSTACK")).size())  
					{
						((Stack)Session.getAttribute("MO_CALLSTACK")).setSize(save_callstack+1);
					    Session.setAttribute("MO_CTX",((Stack)Session.getAttribute("MO_CALLSTACK")).pop());
						Session.setAttribute("MO_TRANSACTION",((Context)Session.getAttribute("MO_CTX")).name);
						Session.setAttribute("MO_CALLER",((Context)Session.getAttribute("MO_CTX")).caller);
					}
					if(save_savestack<((Stack)Session.getAttribute("MO_SAVESTACK")).size())
					{
						((Stack)Session.getAttribute("MO_SAVESTACK")).setSize(save_savestack+1);
						restore();
					}
				}
				if(mo_ctx.option("nocontent"))
				{
				    response.setStatus(204);
				    //Response.Status="204 NO CONTENTS";	//retour sans contenu : n'efface pas la page 
				    throw new  Endexception();
				}

				if ((String) Application.getAttribute("message")!=null && !compare(((String) Application.getAttribute("message")),""))
				{
					if(((Boolean)Session.getAttribute("MO_MESSLU"))==null || !((Boolean)Session.getAttribute("MO_MESSLU")))
					{
						writer.write("<script>alert('"+(String) Application.getAttribute("message")+"');</script>");
						Session.setAttribute("MO_MESSLU",true);
					}
				}
				else
					Session.setAttribute("MO_MESSLU",false);
				
				//if((int)2==(int)1) // && !((Boolean)Session.getAttribute("MO_NOSAVE")))
				if((Integer)(Session.getAttribute("MO_VIAXML"))==1  && !((Boolean)Session.getAttribute("MO_NOSAVE")))
					Session.setAttribute("MO_VIAXML",2);
				else
					Session.setAttribute("MO_VIAXML",0);

		        throw new  Endexception();
		    	//-----------------------------------EXEC page--------------------------------------- EXEC
		      case EXEC :
		        if(((Boolean)Session.getAttribute("MO_INTRANS")) && ((Boolean)Session.getAttribute("MO_VBSERR")))  // on a eu un onerror
		        {
					mo_ctx.geturl();  // sauter le mot suivant 
					break;
				}

		        mo_ctx.lastpage=mpos;   // dernière page exécutée
		        page=mo_ctx.geturl(); // 
				tlog("X",mo_page,false);       
		        MyExecute(page);		//++++
				tlog("","",false);  
				if(compare(verb,"EXECP"))
					Session.setAttribute("MO_NOMPAGE",mo_page+" ");
				else if(((String)Session.getAttribute("MO_NOMPAGE")).indexOf(" ")==-1 && !compare(mo_page,"PGL_COMMAND.JSP"))
					Session.setAttribute("MO_NOMPAGE",mo_page);
				
		        break;
		                
		//-----------------------------CALL trans---------------------------------------------CALL
		      case CALL :     // idem mais on continue en séquence

				mo_ctx.issaved=true;
		        mo_ctx.xpush(new Ct(VERB.CALL,mo_ctx.lastpage));   // on empile lastpage (inutilisé pour CALL) 

		//------------------------------SAVE-------------------------------------------- SAVE
		      case SAVE :     // sauvegarde de l'objet Session dans SAVESTACK
				save();		
		        break;
		        
		//----------------------------RESTORE---------------------------------------------- RESTORE
		      case RESTORE :  // restauration Session
		        restore();
		        break;
		        
		//-----------------------------TRANS transaction--------------------------------------------- TRANS
		      case TRANS :  // appel d'un script enregistré dans le dictionnaire
		//-----------------------------TRANS transaction--------------------------------------------- TRANS
		      case BRANCH :  // appel d'un script RAZ des stacks
		//-----------------------------TRANSFER transaction------------------------------------------ TRANSFER
		      case TRANSFER :  // appel d'un script sans retour    
		        t=mo_ctx.gettrans();   // lecture du nom 
				tlog("T",t,true);
		        //p=TRANS.content(t);  // texte du script dico.dict
		        p=texttrans(t);  // texte du script
		        if(p==null||compare(p,""))
					error("13 Transaction "+t+" inexistante");
		        if(compare(verb,"TRANS"))
					nctx=new Context(p,t,((String)Session.getAttribute("MO_TRANSACTION")));  // initialisation du contexte
				else
					nctx=new Context(p,t,""); 
					
		        if(compare(verb,"TRANS"))
					((Stack)Session.getAttribute("MO_CALLSTACK")).push(mo_ctx);  // empilement de l'ancien contexte
		        else if(compare(verb,"TRANSFER"))
		        {
					nctx.transfer=true;
					nctx.request=mo_ctx.request;
				}
		        Session.setAttribute("MO_CTX",nctx);  // le nouveau contexte est conservé dans la session
		        if(compare(verb,"BRANCH"))
		        {
					Session.setAttribute("MO_SAVESTACK",newArray());//Session.setAttribute("MO_SAVESTACK").name="MO_SAVESTACK");
					Session.setAttribute("MO_CALLSTACK",newArray());//Session("MO_CALLSTACK").name="MO_CALLSTACK";
				}
		        play(nctx);  // on l'exécute en récursif .   Au cas où l'on revient : 
		        if(compare(verb,"BRANCH"))
		        {
					throw new  Endexception();	
				}		
			    if(nctx.request && nctx.transfer)
					((Stack)Session.getAttribute("MO_CALLSTACK")).pop();
			    if(((Stack)Session.getAttribute("MO_CALLSTACK")).size()>0)
					mo_ctx=(Context)((Stack)Session.getAttribute("MO_CALLSTACK")).pop(); // on dépile l'ancien contexte
			    if(mo_ctx.issaved)
			    {
					restore();
					mo_ctx.issaved=false;
		        }       
				Session.setAttribute("MO_TRANSACTION",mo_ctx.name);
				Session.setAttribute("MO_CALLER",mo_ctx.caller);
		        Session.setAttribute("MO_CTX",mo_ctx);    // on le stocke dans la session
			    mo_skipcase=false;
			    continue;
				//-------------------------------CASE condition------------------------------------------- CASE
		      case CASE :    // test
		      
		        if(mo_skipcase ){ // si l'on est déja tombé dans un CASE, on saute jusqu'à un
		          mo_ctx.skip();   // ordre différent de CASE ou OTHER
		          continue;
		        }
		      
		        not=mo_ctx.option("NOT");  // CASE NOT
		        C=mo_ctx.getcond();		  // lecture condition: code retour prédéfini ou utilisateur 
		        if(compare(C,"EMPTY"))
					C="";
		        d=(Hashtable)Session.getAttribute("MO_DROITS");   // ou DR:<code droit>
		        if (compare(C.substring(0,2),"DR:")){
		          if((Integer)d.get(C.substring(3))!=1^not)
		          { // le dictionnaire contient 1 si droit = oui
		            mo_ctx.skip();			//  non autorisé
		            continue;         
		          }
		        }
		        else					// code retour 
		          if((!compare(C,mo_rc))^not){
		            mo_ctx.skip();		// test négatif
		            continue;
		          }
		        mo_ctx.xpush(new Ct(VERB.CASE,0)); // on empile le verbe CASE 
		        break;
		//-------------------------------CASE condition------------------------------------------- CASE
		      case CASE2 :    // test multiple
		      
		        if(mo_skipcase ){ // si l'on est déja tombé dans un CASE, on saute jusqu'à un
		          mo_ctx.skip();   // ordre différent de CASE ou OTHER 
		          continue;
		        }
		        arr=mo_ctx.getlist2();
		        ll=false; 
		        for( i=0;i<arr.size();i++)
		        {
		          if(compare(C,mo_rc))
		          {
					ll=true;
					break;
		          }
		        }  
		        if(!ll)
		        {
		            mo_ctx.skip();		// test négatif
		            continue;
				}
		        mo_ctx.xpush(new Ct(VERB.CASE,0)); // on empile le verbe CASE
		        break;
		        
		 //----------------------------OTHER---------------------------------------------- OTHER
		     case OTHER :			// suit facultivement une série de CASE
									// exécuté sauf si skipcase positionné
		        if(mo_skipcase ){
		          mo_ctx.skip();
		          continue;
		        } 
		        mo_ctx.xpush(new Ct(VERB.ANY,0)); 
		        break;
		          
		//-----------------------------REPEAT instr--------------------------------------------- REPEAT
		      case REPEAT :		// boucle (arrêtée par JUMP ou BREAK)
		      
		        mo_ctx.xpush(new Ct(VERB.REPEAT,mpos)); // on mémorise la position du début d'instruction REPEAT 
		        break;
		        		        
		//-------------------------------ON cond instr------------------------------------------- ON
		//-------------------------------CONTINUE------------------------------------------- CONTINUE
		      case CONTINUE :          // NOOP
		        break;
		    	//-------------------------------EVAL script sans mise résultat dans MO_RETURN---- §  §
		      case EVAL2:     // évaluation javascript
		      
		        mo_ctx.ptr=mpos;  // on revient en début d'instruction
		        txt=mo_ctx.geteval();   // lecture du script
		        try{
		          eval(txt);
		        }  
		        catch(Exception e){
		          error("14 erreur dans l'exécution du script: \n"+txt);
		        }
			    mo_skipcase=false;
			    continue;
		//-------------------------------EVAL script avec mise résultat dans MO_RETURN----- EVAL
		      case EVAL :
		        txt=mo_ctx.geteval();   // lecture du script
		        x="";
		        try{
		          x=(String)eval(txt);
		        }  
		        catch(Exception e){
		          error("15 erreur dans l'exécution du script: \n"+txt);
		        }
		        if(x!=null)
		        {
					try {Session.setAttribute("MO_RETURN",x.toString());} // résultat dans code retour
					catch(Exception e){Session.setAttribute("MO_RETURN","");}
				}
		        break;
		    	//----------------------------------READTRANS fic---------------------------------------- READTRANS
		      case READTRANS :  // lecture d'un fichier de transactions     
		        file=Application.getRealPath(mo_ctx.geturl());   // nom du fichier
		        readtrans(file);
		        break;


		//---------------------------------------------------------------------------------------------
		//---------------------------------------------- CODES RETOUR EGALEMENT TRAITES COMME ORDRES
		//---------------------------------------------------------------------------------------------
		      case ERROR:
		        if(mo_ctx.lastaff<0)
		        {
		          Session.setAttribute("MO_RETURN","ERROR");
		          return;		// retour script supérieur
		        }
		        mo_ctx.ptr=mo_ctx.lastaff;
		        mo_ctx.stack=(Stack)mo_ctx.laststack.clone();
		        break;

		      case AGAIN :
		        if(mo_ctx.lastpage<0)
		        {
		          Session.setAttribute("MO_RETURN","AGAIN");
		          return;		// retour script supérieur
		        }
		        mo_ctx.ptr=mo_ctx.lastpage;
		        break;          
		      
		      case LOOP:
		      case BREAK:
		        bfound=false;      
	            while((ct=mo_ctx.xpop())!=null )
		        {
		          if(ct.verb==VERB.REPEAT)
		          {
		            bfound=true;
		            mo_ctx.ptr=ct.data;
		            if(compare(verb,"BREAK"))
		              mo_ctx.skip();
		            break;
		          }  
		        }
		        if(!bfound)
		        {
		          Session.setAttribute("MO_RETURN",verb);
		          return;
		        }
		        break;
				
			 case MODAL :
				 Session.setAttribute("MO_MODAL","YES");
				 break;
			 case ENDMODAL :
			    if(!compare(Session.getAttribute("MO_MODAL"),"NO"))
			    {
					if(((Stack)Session.getAttribute("MO_CALLSTACK")).size()>0)
					{
						Session.setAttribute("MO_CTX",((Stack)Session.getAttribute("MO_CALLSTACK")).pop()); 
						if(((Context)Session.getAttribute("MO_CTX")).issaved)
						{
							restore();
							((Context)Session.getAttribute("MO_CTX")).issaved=false;
						}
						Session.setAttribute("MO_TRANSACTION",((Context)Session.getAttribute("MO_CTX")).name);
						Session.setAttribute("MO_CALLER",((Context)Session.getAttribute("MO_CTX")).caller);
					}		    
					Session.setAttribute("MO_MODAL","NO");
		/*
			<script> 
				document.cookie = "modal=NO;path="+escape("/");
//				document.cookie = "lastnum=<%=Session('MO_NPAGEOUT')%>;path="+escape("/");
				document.cookie = "lastnum=0;path="+escape("/");
				document.cookie = "result=<%=escape(Session.getAttribute('MO_RESULT'))%>;path="+escape("/");
				window.close();
			</script>
		*/
				    Session.setAttribute("MO_NPAGEOUT",null);
					doEnd();
				}
				continue;

		      case LANGUAGE :  // langue
		        Session.setAttribute("MO_LANGUE",mo_ctx.getword().toLowerCase());		// code langue
				break;
		        
		      default : // instruction erronnée
		      
		        error("21 erreur script verbe inconnu: "+verb);
		       // writer.write("erreur script pos: "+mo_ctx.ptr+" verbe: "+verb);
		       // throw new  Endexception();
		    }
		    
		    mo_skipcase=false;// pour CASE+skip on ne passe pas par là
		    mo_rc="";
		    if(!compare((r=(String)Session.getAttribute("MO_REQUEST")),""))			// retour Server.exec
		    {
				nctx=new Context(r,"request","");
				nctx.request=true;   
				((Stack)(Session.getAttribute("MO_CALLSTACK"))).push(mo_ctx);  // empilement de l'ancien contexte
				Session.setAttribute("MO_CTX",nctx);  // le nouveau contexte est conservé dans la session
				Session.setAttribute("MO_REQUEST","");
				play(nctx);  // on l'exécute en récursif .   Au cas où l'on revient : 
				if(nctx.request && nctx.transfer)
					((Stack)Session.getAttribute("MO_CALLSTACK")).pop();
				mo_ctx=(Context)((Stack)Session.getAttribute("MO_CALLSTACK")).pop(); // on dépile l'ancien contexte
		      Session.setAttribute("MO_CTX",mo_ctx);    // on le stocke dans la session
		    }
		 }
	}

}