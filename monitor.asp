<%@ Language=JavaScript %>
<!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" -->  
<%
//e= main program of the application
//e= monitoring of the ASP pages execution
//e= miscellaneous functions : management of COM components , context's backup,...

var bfirst=(Request("mode").Item!=null);
var tmp=Request("aspsessionid").Item;
var bmodal=(tmp!=null);
if(bmodal && tmp!=getAspSessionID())
{
	%>	
	<script>
		alert("close the second browser and retry");	
		window.close();
	</script>	
	<%	
	    Response.End();
}

if(   Session("MO_SAVESTACK")==null && !bfirst)
{
	%>
		<script>alert("your session is ended");</script>
	<%		Response.End ();
}
//*************************************************************************************
//*************************************************************************************



var xxxtran,stname,stline,stinstr;
xxxtran=0;
Session("MO_ERRMESS")="";

var mo_caprockvect=true;		 // use caprock.vector in place of Array(single thread)
var mo_persist_connection=false;  // persist ORACLE  connection  
var TRANS;						 // transactions dictionary
var save_callstack, save_savestack;

// LOG ****************
var nasterlog="N";  // logging parameters
var nasterloglognaster="nasterlog3.lognaster"; // progid 
var tlevel=Application("MO_TLOGLEVEL"); // from global.asa : 
		// 0=none   1=errors  2=1+start/end sessions 3=2+perf  4=all

var domDocumentProgId = "msxml2.freethreadeddomdocument"; //MSXML3
var templateProgId = "MSXML2.XSLTemplate" //MSXML3

var objXMLS;    //www


	var sMO_NBINST=0;		// DEBUG stuff
	var sMO_INSTANCES="";  
	var sMO_STACK=0;
	var sMO_NBINST2=0;		

if(bfirst)	// 1st call
{
	var inutile=Server.CreateObject("xmlservice.document2");  
	var o=Server.CreateObject("caprock.vector");  // DEBUG
	if(o.noinst)   // debug version of CAPROCK    
	{
		Session("MO_VECTOR")=o;			
		Session("MO_DICT")=Server.CreateObject("caprock.dictionary"); 
		Session("MO_VECTOR").name="$";  
	}
}
else											
{
	if(Session("MO_VECTOR"))		// DEBUG
	{
		var sMO_NBINST=Session("MO_VECTOR").nbinst;
		var sMO_INSTANCES=Session("MO_VECTOR").instances;
		var sMO_STACK=Session("MO_CALLSTACK").length;
		var sMO_NBINST2=Session("MO_DICT").nbinst;
	}
}
var XSession=newArray2 ();


// signal load end for timing purpose
if(Request("endsend").Item=="yes" || Request("endsend").Item=="YES")	
{
	//mo_userid=Request.ServerVariables("REMOTE_ADDR");
	mo_user=Session("EX_UTILISATEUR");
	mo_time=Session("MO_TIME");
	mo_eventtype="S";
	mo_name="SEND";
	mo_tlog=Session("MO_TLOG");
	if(Session("MO_ALERT")==0 && mo_time!=-1)  
	{
		tlog();
		tlogperf(Request("timedif").Item);
		Session("MO_TIME")=-1;
	}
	else   // alert box prevents from perf evaluation
	{
		mo_time=-1;
		tlog();
		tlogperf(-1);
	}
	if(Request("endsend").Item=="yes")	
		Response.Status="204 NO CONTENTS";	
	else
		Response.Write("<"+"script>window.close();<"+"/script>");

    ResponseEnd();
}
Session("MO_ALERT")=0; 
// the client signals end modal window
// due to a bug in IE window.open that causes sessions swap
// also called by window.asp if the user closes the window using the system icon

if(Request("endmodal").Item=="yes")
{ 
				// restore stacks s
	sc=Session("MO_SAVE_CALLSTACK");
	if(sc!=null && sc<Session("MO_CALLSTACK").length)
	{
		setcount(Session("MO_CALLSTACK"),sc+1);
	    Session("MO_CTX")=pop(Session("MO_CALLSTACK"));
		Session("MO_TRANSACTION") =Session("MO_CTX")("name");
		Session("MO_CALLER") =Session("MO_CTX")("caller");
	}
	ss=Session("MO_SAVE_SAVESTACK");
	if(ss<Session("MO_SAVESTACK").length)
	{
		setcount(Session("MO_SAVESTACK"),ss+1);
		restore();
	}
	Session('MO_NPAGEOUT')=null;
	Session("MO_MODAL")="NO";
    Response.Status="204 NO CONTENTS";
    ResponseEnd();
}

%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"> -->
<HEAD><meta http-equiv="expires" content="-1"></HEAD>

<%

Session("MO_NOMPAGE")="";
var mo_rc="";
Session("F_InControl") = true;
Session("F_InPage") = true;
site=Application("MO_SITE");
//-----------------------------------------------------------------------------
//errors processing : moniteur.asp used for 500-100
//-----------------------------------------------------------------------------
err=Server.GetLastError();
if(err.Number!=0)
  error("1 GetLastError");


bForm=(Request.ServerVariables("REQUEST_METHOD")=="POST"); 
							// Session("MO_REQUEST")= script 
req=Session("MO_REQUEST");	
Session("MO_REQUEST")="";   

//------------------------------------------------------------------------------
//  initializations
//------------------------------------------------------------------------------
//  Session("MO_REQUEST")   request script
//  Session("MO_DATASOURCES")	data source
//  Session("MO_SAVESTACK")  context stack
//  Session("MO_CALLTACK")  script calls stack
//  Session("MO_TRANS")   transactions dictionary
//  Session("MO_SESSTRANS") flag: transactions dictionary in session  object(instead of application object)
//  Session("RFORM")  form backup
//  Session("MO_XMLS")  wrapper xmlDom
//  Session("MO_XSL")  default style sheet 
//  Session("MO_XSL2")  xmldom object
//  Session("MO_RETURN")  return code
//  Session("MO_LASTRETURN") last generated return code
//  Session("MO_RESULT")	result code
//  Session("MO_COMMANDES")	command bar definition
//  Session("MO_ACTION")  performance param
//  Session("MO_INTRANS")  oracle transaction pending
//  Session("MO_NPAGEOUT") auto-incremented page number
//  Session("MO_APPLICATION") application name
//  Session("MO_TRANSACTION") transaction name
//  Session("MO_NOSAVE")	flag: don't save form's datas
//  Session("MO_CLEAR")		flag:  clear xmlservice object before use
//  Session("MO_LANGUE")   language
//  Session("MO_USER")	user id
//  Session("EX_UTILISATEUR") user login
//  Session("MO_DROITS")    user rights dictionary
//  Session("MO_DICWR)      data access wrappers instances dictionary
//  Session("MO_HASERROR")  flag: data error
//  Session("MO_WRNAME")    data wrapper DLL name
//  Session("MO_MENU")		current menu name
//  Session("MO_TITLE")="";   window title
//  Session("MO_NEXTDATE")  wake-up date time
//  Session("MO_NEXTCOMMENT") wake-up comment
//	Session("MO_VBSERR")	flag: VBScript error
//	Session("MO_VIAXML")	flag: use xmlservice object for next request
//  Session("MO_DOCTYPE")    header
//  Session("MO_INSTANCE")    oracle instance
//  Session("MO_SCHEMA")	 oracle schema
//  Session("MO_PASSWORD")    oracle password
//  Session("MO_STARTMENU") staring menu transaction 
//  Session("MO_NOINCR")	flag: inhibits page sequence control

//  MO_TRANS  transactions dictionary
//  MO_FILES  files dictionary 

if(   Session("MO_SAVESTACK")==null
   || bfirst
   || req.toUpperCase()=="INIT"
   || Request("init").Item=="yes")
{
  Session("MO_DOCTYPE")='<HEAD><meta http-equiv="expires" content="-1"></HEAD>';

  Session("MO_XSL")		=Server.CreateObject(templateProgId);
  Session("MO_XSL2")	=Server.CreateObject(templateProgId);
  temp=Server.CreateObject(domDocumentProgId);
  temp.async=false;
  nxsl="SERVICE-1.0-TO-XHTML-1.0-FRAGMENT.XSL";
  nxsl=Server.MapPath("../xsl/"+nxsl);
  r=temp.load(nxsl);
  if(!r)
    error("2 échec load XSL: "+ nxsl);
  Session("MO_XSL").stylesheet = temp;
  temp=null;
  
  Session("MO_SAVESTACK")			=newArray();
  Session("MO_CALLSTACK")			=newArray();
  Session("MO_DATASOURCES")			=newArray2();
  Session("MO_RETURN")				="";
  Session("MO_ACTION")				=""; 
  Session("MO_RESULT")				=null;
  Session("MO_COMMANDES")			="";	
  Session("MO_INTRANS")				=false; 
  Session("MO_LANGUE")				="undefined"; //"fr"; //"en";
  Session("MO_NOSAVE")				=false; 
  Session("MO_CLEAR")				=false; 
  Session("MO_LASTRETURN")			="";  
  Session("MO_APPLICATION")			="NASTER";
  Session("MO_TRANSACTION")			="NASTER";
  Session("MO_CALLER")				="";
  Session("MO_USER")				="";
  Session("EX_UTILISATEUR")			="";
  Session("EX_UTILISATEUR2")		="";
  Session("MO_HASERROR")			=false;
  Session("MO_NOALERTERROR")		=false;
  Session("MO_WRNAME")				="PLWRAPPER";
  Session("MO_MENU")				="NASTER";
  Session("MO_OLD_MENU")			="";
  Session("MEM_GL_BLINKTIME")		=0;
  Session("MEM_GL_WIDTH")			=0;
  Session("MO_DROITS")				=newArray2()
  Session("MO_DICWR")				=newArray2();
  Session("MO_NEXTDATE")			=null;
  Session("MO_NEXTCOMMENT")			="";
  Session("MO_ISLOG")				=false;
  tmp								=Request.ServerVariables("SERVER_NAME").Item;
  i									=tmp.indexOf(".");
  if(i>=0)
	tmp								=tmp.substring(0,i);
  tmp								=tmp.substr(0,10);
  Session("MO_SERVER")				=tmp+".mcsfr.net";		
  if (Application("MO_SERVER")==null) 
	Application("MO_SERVER")		=Session("MO_SERVER");
  Session("MO_VBSERR")				=false;
  Session("MO_VIAXML")				=0;     
  Session("MO_SENT_TRANS")			="";
  Session("MO_INSTANCE")			="";
  Session("MO_SCHEMA")				="";
  Session("MO_PASSWORD")			="";
  Session("MO_APPLI")				="";
  Session("MO_MESSLU")				=false;
  Session("MO_MESSEXT")				=true;
  Session("MO_STARTMENU")			="TRANS MENU;"; 
  Session("MO_NOINCR")              =false;
  
  tmp=Request("instance").Item;	
  if(tmp!=null)
	Session("MO_INSTANCE")	=tmp; 
    
  tmp=Request("schema").Item;
  if(tmp!=null)
	Session("MO_SCHEMA")	=tmp; 
	
  tmp=Request("password").Item;
  if(tmp!=null)
	Session("MO_PASSWORD")	=tmp; 
	
  tmp=Request("appli").Item;
  if(tmp!=null)			
	Session("MO_APPLI")		=tmp;
  else
	Session("MO_APPLI")		="TSTNS3"; 
	
  tmp=Request("wrname").Item;
  if(tmp!=null)			
	Session("MO_WRNAME")	=tmp;
	
  tmp=Request("mode").Item;
  if(tmp==null || tmp=="test")		
  {
	site=Application("MO_SITE");
	Server.Execute ("/"+site+"/asp/mo_initappl.asp");		
  }
  
  tmp=Request("flavor").Item;
  if(tmp!=null)			// appli = CTX	
	Session("MO_FLAVOR")=tmp; 
  else
	Session("MO_FLAVOR")=""; 

      
  var tmptrans=Request("trans").Item;
  if(tmptrans!=null)	
	Session("MO_STARTMENU")="TRANS "+tmptrans+";";
      
  tmp=Request("startmenu").Item;
  if(tmp!=null)			// appli = CTX	
	Session("MO_STARTMENU")=tmp; 

  tmp=Request("noincr").Item;
  if(tmp!=null)			// appli = CTX	
	Session("MO_NOINCR")=true; 

  if(Session("MO_NPAGEOUT")==null)
    Session("MO_NPAGEOUT")="0";
  if(req==null)
  {
    Session("MO_REQUEST")="";
    req="";
  }
  TRANS= MO_TRANS;
  Session("MO_SESSTRANS")=false;    
  req=Request.QueryString("mo_request").Item;
  if(req==null)  
  {
	if(tmptrans==null)
	   tr="NASTER";
	else
	   tr=tmptrans;
	req="TRANS "+tr+";";
  }
  if(nasterlog=="P")
	Session("MO_TLOG")=Server.CreateObject(nasterloglognaster);			// persistent
  else
	Session("MO_TLOG")=null;
  mo_tlog=Session("MO_TLOG");
  
  Session("MO_TIME")=0;		                           
  Session("MO_MODAL")="NO";
}
else
  
//------------------------------------------------------------------------------
//  called by SUBMIT (POST or GET) from a displayed page
//------------------------------------------------------------------------------
{
  if(Session("MO_TRANS")==null)
	TRANS=MO_TRANS;
  else
	TRANS=Session("MO_TRANS");

//------------------------ NAVIGATION CONTROL -------------------------
// prevents user from backward-forward navigation thru browser

  if(REQF("noincr")==null && !Session("MO_NOINCR"))  
  {
    npagein= REQF("MO_NPAGE");  // page number stored in HTML
    npageout= Session("MO_NPAGEOUT");        // current number in Session object
    if(npageout==null)
		npageout=npagein;
		
    if (npagein!=null && npagein!="" && npagein!= npageout){			
      Response.Status="204 NO CONTENTS";
      Session("F_InControl") = false;
	  Session("F_InPage") = false;
      ResponseEnd();
    } 
    
    n=eval(npageout)+1;					
    Session("MO_NPAGEOUT")=n.toString();
  }
  r=REQF("MO_RESULT");
  if(r!=null)						
	Session("MO_RESULT")=r;
  r=REQF("MO_REQUEST");				
  if(r!="" && r!=null)
    req=r;
  r=REQF("MO_RETURN");
  if(r!=null && r!="")
  {
    Session("MO_RETURN")= r;  
    Session("MO_ACTION")= r;  
  }
  else
    Session("MO_ACTION")= "";  
  
    
//-------------------- SAVE the FORM -----------------------------------  
  if(bForm)
    rq=Request.Form;
  else
    rq=Request.QueryString;
    
//  MEM_.. variables are automatically stored in session
  for(i=1;i<=rq.Count;i++)
  {
    k=rq.Key(i);
    if(k.substr(0,4)=="MEM_")
		Session(k)=rq.Item(i).Item;
  }
    
// store in dictionary
  dico = newArray2();
  for(i=1;i<=rq.Count;i++){
    k=rq.Key(i).toUpperCase();
    try{
    dico(k)=rq.Item(i).Item;
    }catch(e){}
  }
  Session("RFORM")=dico;
  dico=null;
//-------------------------- restore XMLS object ----------
  restorexmls(); 

//							Save form in xmlservice object
  if(bForm && !Session("MO_NOSAVE")) 
  {
	w=XmlServiceFactory2("MO_XMLS");
	if(w!=null)
	{
	   try{
	    l=w.Service.ElementList;
		if(Session("MO_CLEAR"))
			for (i=1;i<=l.Count;i++)
				l(i).SingleValue="";		
		}catch(e){Session("MO_ERROBJ")=e;error("4 restau XMLS");}		
		
		for (i=1;i<= rq.count;i++)
		{
		  name=rq.Key(i).toUpperCase();
		  if(name.substr(0,4)=="MEM_")
		  {
				Session(name)=rq.Item(i).Item;
				name=name.substr(4);
		  }
		  try{
			 el=l(name);
	         dl=el.DataList;
		     dl.Clear();
		     it=rq.Item(i);
		     for (j=1;j<= it.Count;j++)
	         {
		       value=it(j);
		       if(el.DataType=="integer" || el.DataType=="double" ) //|| el.DataType=="date"))
					if(rtrim(value)=="")					
						value=0;
					else
						value=asnum(value);
		       if(j==1)    
		         dl.Item(j).Value=value;
		       else
		         dl.Add().Value=value;	         
		     }
		  }catch(e){};
		}
		l=null;
		el=null;
		dl=null;
		it=null;
	}
	w=null;
  }

  Session("MO_CLEAR")=false;
}
mo_wrname=Session("MO_WRNAME");

//------------------------------------------------------------------------------
//    SCRIPTS PROCESSING
//------------------------------------------------------------------------------
// Session(MO_CTX)= current script context
while(true)														// loop on starting menu
{
	save_callstack=Session("MO_CALLSTACK").length;
	save_savestack=Session("MO_SAVESTACK").length;
	if(req !=null && req!="")
	{    // if request exist
		 // for a modal window, save stacks levels in case of uncontroled closing
	  if(bmodal)
	  {
		Session("MO_SAVE_CALLSTACK")=save_callstack;
		Session("MO_SAVE_SAVESTACK")=save_savestack;
		bmodal=false;
	  }
	// save current script if required
	  if(Session("MO_CTX")!=null)
	  {
	    ctx=Session("MO_CTX");
	    push(Session("MO_CALLSTACK"),ctx); // push(arr,x) remplace arr.push(x) (bug)
	  }
	// initialize script
	  ctx=newArray2();
	  inittrans(ctx,req,"request","");
	  ctx("request")=true;
	  Session("MO_CTX")=ctx;
	}   

	// initialize jumps and imbrication
	mo_imbricate=false;	
	mo_jump=0;			
	mo_label="";        

	if(Session("MO_ISLOG"))
		mo_tlog=Session("MO_TLOG");
	else
		mo_tlog=null;
		
	mo_eventtype="";
	mo_name="";
	mo_time=0;
	mo_user=Session("EX_UTILISATEUR");
	mo_page="";
	
	//loop on stacked scripts
	while(true)				
	{
	  play(Session("MO_CTX"));   // execute script
	  tlog("T","/"+Session("MO_CTX")("name")+"/",true);
	  if(Session("MO_CALLSTACK").length==0)
	    break;
	  if(Session("MO_CTX")("request") && Session("MO_CTX")("transfer"))
	  {
		  if(Session("MO_CALLSTACK").length==0)
			break;
		  pop(Session("MO_CALLSTACK"));
	  }
	  Session("MO_CTX")=pop(Session("MO_CALLSTACK")); // script suivant dans la pile
	  if(Session("MO_CTX")("issaved"))
	  {
		restore();
        Session("MO_CTX")("issaved")=false;
      }
      Session("MO_TRANSACTION") =Session("MO_CTX")("name");
	  Session("MO_CALLER") =Session("MO_CTX")("caller");
	}
	Session("MO_CTX")=null;
	req=Session("MO_STARTMENU");	
    Session("MO_RETURN")="";
}	
// reinitialize when called and session not closed
Session("MO_SAVESTACK")=null;

Session.Abandon();

//---------------- FUNCTIONS ---------------------------------------

//--------------------------------------------------------------------
function play(mo_ctx)
{							//		process script
//--------------------------------------------------------------------


// mo_ctx("name") = nom
// mo_ctx("caller")= caller transaction
// mo_ctx("prg") = script text
// mo_ctx("ptr") = instruction pointer
// mo_ctx("stack") = execution stack
// mo_ctx("laststack") = stack when last page was sent
// mo_ctx("lastpage") = last called page
// mo_ctx("lastaff") = last displayed page
// mo_ctx("ons") = ON criterias
// mo_ctx("labels") = labels
// mo_ctx("issaved")= flag: context saved
// mo_ctx("request")
// mo_ctx("transfer")= transfer in progress

// readtrans()  read transactions files (scripts)
// inittrans()  initialize execution context
// getlabels()   get labels
// dojump()     jump
// dojumpnext()     
// dojumpprev()     
// xpush()       push
// xpop()        pop
// getword()    read next token 
// getverb()    read next verb
// gettrans()   read transaction name
// geturl()     read url
// getcond()    read  condition code
// geteval()    read javascript text for eval
// getstring()  read string
// getlist()	read arg list
// skip()		skip instruction
// skipcom()	skip comment
// option()     read option
// restore()    restore stack from Session("MO_SAVESTACK") 
// header()		 menu
// footer()		 commands
// dobegintrans()  open database transaction
// docommit()    commit
// dorollback () rollback
// isintrans()	test oracle transaction pending
// clone()	copyarray()	push() pop() clone/copy/push/pop array
// save() save context
// doEnd() response.end


  var lab,mo_skipcase,mo_more,tmp,nptr,bfound,verb,ct,k,s,i,p,t,arr,c,d,txt,m,r;
  var x,file,a,cnx,cmd,src,ds,a2,nctx,isin,wrapp,ors,w,oxmls,name,commandes,lang,langa,bcon2;
  var ltr,cnt,plang,wrp;
  //var mo_rc;
  Session("MO_TRANSACTION") =mo_ctx("name");
  Session("MO_CALLER") =mo_ctx("caller");

  if(mo_imbricate)
  {
    mo_ctx("ptr")=mo_ctx("lastpage");  
    mo_imbricate=false;          
  }
  
  if(mo_ctx("ptr")==0)				
    getlabels(mo_ctx);
    
  if((lab=mo_label)!="")  // process jumps
  {							
    switch(mo_jump){       
      case 0:              
        if(!dojump(mo_ctx,lab))  
          return;
        break;                
      case -1:						// -1 JUMPPREV
        if(!dojumpprev(mo_ctx,lab))
          return;
        break;
      case 1:						//  1 JUMPNEXT
        if(!dojumpnext(mo_ctx,lab))
          return;
        break;
    }
    mo_label="";         
  }
    
  mo_skipcase=false;   // flag : skip CASE after CASE found
  mo_more=true;        // flag : more loop
  
//-------------------------- instruction loop -------------------------------------------
  
  while(mo_more)
  {
    
    //------------------------process return codes-------------------------
    
    tmp=Session("MO_RETURN").toUpperCase( ); 
    if(tmp!="")
    {	
      if (tmp!="ERROR")
		Session("MO_LASTRETURN")=tmp;    
      i=tmp.length;
      while(--i>0)		// trim blanks and ;
        if((c=tmp.charAt(i))!=" " && c!=";")
          break;
      tmp=tmp.substring(0,i+1);
      mo_rc=tmp;
   // look for ON clause
      p=mo_ctx("ons")(mo_rc);
      if(p!=null)
      {
		nctx=newArray2();   
		inittrans(nctx,p,"ONREQUEST",Session("MO_TRANSACTION"));  
		nctx("request")=true;
		Session("MO_CTX")=nctx;  			
		Session("MO_RETURN")=""; 
        push(Session("MO_CALLSTACK"),mo_ctx);  
		play(nctx);  // recurse script 
	    if(Session("MO_CALLSTACK").length>0)
           mo_ctx=pop(Session("MO_CALLSTACK")); 
      } 
      else
      {
        switch (mo_rc) 
        {
          case "ERROR":		// go to last displayed page if defined
            if(mo_ctx("lastaff")<0)
               return;	
            mo_ctx("ptr")=mo_ctx("lastaff");
            mo_ctx("stack")=clone(mo_ctx("laststack"));
            break;
          case "AGAIN" :     // reexecute last page
            if(mo_ctx("lastpage")<0) 
              return;
            mo_ctx("ptr")=mo_ctx("lastpage");
            break;          
          case "LOOP":	// go to REPEAT
          case "BREAK":     // exit loop
                // look for a REPEAT
            bfound=false;
            while((ct=xpop(mo_ctx))!=null )
            {
              if(ct("verb")=="REPEAT")
              {
                mo_ctx("ptr")=ct("data");  // position on loop start
                if(mo_rc=="BREAK")   
                  skip(mo_ctx);    // skip REPEAT
                bfound=true;
                break;
              }  
            }
            if(!bfound)
              return;			
            break;
        
          case "ERRSYS" :
            error("6 erreur systeme");
            break;
          case "ABORT" :  
          case "RETURN" :
            mo_more=false;
            if(mo_ctx("name")!="request")
            {
              Session("MO_RETURN")="";  
              mo_rc="";
            }
            continue;
          case "ABANDON" :  // stop session
			doAbandonAppli();
		  case "ENDMODAL" :  // end modal window
		    if(Session("MO_MODAL")!="NO")
		    {
				if(Session("MO_CALLSTACK").length>0)
				{
					Session("MO_CTX")=pop(Session("MO_CALLSTACK")); 
					if(Session("MO_CTX")("issaved"))
					{
						restore();
					    Session("MO_CTX")("issaved")=false;
					}
					Session("MO_TRANSACTION") =Session("MO_CTX")("name");
					Session("MO_CALLER") =Session("MO_CTX")("caller");
				}		    
				Session("MO_MODAL")="NO";
%>
	<script>
		document.cookie = "modal=NO;path="+escape("/");
		document.cookie = "result=<%=escape(Session('MO_RESULT'))%>;path="+escape("/");
		window.close();
	</script>
<%		    
				Session('MO_NPAGEOUT')=null;
				doEnd();
			}
			break;
			
		  default :
			c=mo_rc.substr(0,1);
			if(c==">")			//	  (TRANSFER)
			{
				t=mo_rc.substr(1);
				Session("MO_MENU")=t;												
				tlog("T",t,true);
				p=texttrans(t);  
				if(p==null)
					error("7 Transaction "+t+" inexistante");
				nctx=newArray2();   // create execution context
				inittrans(nctx,p,t,""); // initialize context
				nctx("transfer")=true;
				nctx("request")=mo_ctx("request");
				Session("MO_CTX")=nctx;  // save new context in session				
				Session("MO_RETURN")=""; 
				play(nctx);  // recurse
				if(nctx("request"))
					pop(Session("MO_CALLSTACK"));
				if(Session("MO_CALLSTACK").length>0)
					mo_ctx=pop(Session("MO_CALLSTACK")); 
				if(mo_ctx("issaved"))
				{
					restore();
					mo_ctx("issaved")=false;
				}       
				Session("MO_TRANSACTION") =mo_ctx("name");
				Session("MO_CALLER") =mo_ctx("caller");
				Session("MO_CTX")=mo_ctx;    
			}		    
			else if(c=="@")			//	   (BRANCH)
			{
				t=mo_rc.substr(1);
				Session("MO_MENU")=t;												
				tlog("T",t,true);
				p=texttrans(t);  // script text
				if(p==null)
					error("8 Transaction "+t+" inexistante");
				nctx=newArray2();   
				inittrans(nctx,p,t,""); 
				Session("MO_CTX")=nctx;  
				Session("MO_SAVESTACK")=newArray();
				Session("MO_CALLSTACK")=newArray();
				Session("MO_RETURN")=""; 
				play(nctx);   
				if(nctx("request"))
					pop(Session("MO_CALLSTACK"));
				if(Session("MO_CALLSTACK").length>0)
					mo_ctx=pop(Session("MO_CALLSTACK")); 
				if(mo_ctx("issaved"))
				{
					restore();
					mo_ctx("issaved")=false;
				}       
				Session("MO_TRANSACTION") =mo_ctx("name");
				Session("MO_CALLER") =mo_ctx("caller");
				Session("MO_CTX")=mo_ctx;    
			}		    
			else if(c=="=")			//	   (CALL)
			{
				t=mo_rc.substr(1);
				Session("MO_OLD_MENU")=Session("MO_MENU");												
				Session("MO_MENU")=t;												
				tlog("T",t,true);

				mo_ctx("issaved")=true;  // CALL
		        xpush(mo_ctx,"CALL",mo_ctx("lastpage"));   // CALL
		        Session("MO_LASTRETURN")="";    // CALL
		        Session("REINIT")=1;				// used by pages that process Request.Form
				save();		// CALL


				p=texttrans(t);  
				if(p==null)
					error("9 Transaction "+t+" inexistante");
				nctx=newArray2();   
				inittrans(nctx,p,t,Session("MO_TRANSACTION"));  
				Session("MO_CTX")=nctx;  
				
				push(Session("MO_CALLSTACK"),mo_ctx);  // CALL

				Session("MO_RETURN")=""; 
				play(nctx);   
				if(Session("MO_CALLSTACK").length>0)
					mo_ctx=pop(Session("MO_CALLSTACK")); 
				if(mo_ctx("issaved"))
				{
					restore();
					mo_ctx("issaved")=false;
				}       
				Session("MO_TRANSACTION") =mo_ctx("name");
				Session("MO_CALLER") =mo_ctx("caller");
				Session("MO_CTX")=mo_ctx;    
			}		    
       }
      }
    }
    Session("MO_RETURN")="";  
    
    mpos=mo_ctx("ptr");		// before instrution
    
    if((verb =getverb(mo_ctx))=="") // read next token 
      break;						

    if(verb=="*")					// comment
    {
      skipcom(mo_ctx);
      continue;
    }
	   
    if(verb==";")
    {                 //  INSTRUCTION END
        cont=true;				// lookup stack	
        while(cont && (ct=xpop(mo_ctx))!=null )
        {
          switch(ct("verb"))
          {
            case "BLOCK" :			//  opening "("
              xpush(mo_ctx,"BLOCK",""); 
              cont=false;         
              break;  
            case "CASE" :  // skip next CASEs
              mo_skipcase=true;
              break;
            case "REPEAT" :  // go to loop start
              mo_ctx("ptr")=ct("data");
              mo_skipcase=false;  // any token different from CASE reinitialize case flag
              cont=false;
              break;
            case "IMBRICATE" : 
              mo_ctx("ptr")=ct("data");
              if(mo_ctx("ptr")<0){  
                mo_imbricate=true; 
                return;
              }					// no BREAK
            case "CALL" :
              mo_skipcase=false;
              break;
          }    
        } 
        continue;               
    }      

    if(verb==")")
    {              //  pop "("
      xpop(mo_ctx);
      continue;
    }
        
    if(verb=="(")
    {               // push "("
      mo_skipcase=false;
      xpush(mo_ctx,"BLOCK","");
      continue;
    }
    
    bcon2=false;
//----------------------------------------------------------------------------------
//------------------------------------ VERBS --------------------------------------
//----------------------------------------------------------------------------------
//---------------------------ABANDON -----------------------------------
    switch (verb) 
    {
      case "ABANDON" :  // stop session
		doAbandonAppli();
     case "ABORT" :  // stop transaction 
     case "RETURN" :
        mo_more=false;
        if(mo_ctx("name")!="request")
          Session("MO_RETURN")="";   
        continue;
//---------------------------ALERT "message"-----------------------------------
	  case "ALERT":
		c=getstring(mo_ctx);
		Response.Write("<script>alert(unescape('"+escape(c)+"'));</script>");
		break;
//---------------------------TITLE "title"-----------------------------------
	  case "TITLE":
		c=getstring(mo_ctx);
	    lang=getword(mo_ctx);
	    langa=Session("MO_LANGUE").toUpperCase();
	    if(lang==";")
			mo_ctx("ptr")--;
		if(lang==";" && langa=="UNDEFINED" || langa==lang)
		{
			Response.Write("<html><head><title>NASTER V1.0 "+c+"</title></head>");
			Session("MO_TITLE")=c;
		}
		break;
//---------------------------WRNAME <name>------- progID----------------------------
	  case "WRNAME":
	    mo_wrname=getword(mo_ctx);
	    Session("MO_WRNAME")=mo_wrname;
		break;
//---------------------------DEBUG-------------------------------------------
	  case "BP":
		break;
//---------------------------TLOG ON/OFF/START/END log performances into tlog table-----------
	  case "TLOG":   
		c=getword(mo_ctx);
		if(nasterlog=="N")
			break;								
		if(c=="ON")		// start
		{
			if(nasterlog=="Y")  // non persistent
			{			
				mo_tlog=Server.CreateObject(nasterloglognaster);
				Session("MO_TLOG")=false; 
			}
			else
				mo_tlog=Session("MO_TLOG");   // null=no logging
			tlog("O","TLOG ON******",true);	
			Session("MO_ISLOG")=true;
		}
		else if(c=="START")		// log user
		{
			r=mo_tlog;			
			if(!mo_tlog)
				mo_tlog=Server.CreateObject(nasterloglognaster);
			mo_user=Session("EX_UTILISATEUR");
			tlog("D","DEBUT SESSION ******",true);	
			mo_tlog=r;
		}
		else if(c=="END")		// end session
		{
			r=mo_tlog;			
			if(!mo_tlog)
				mo_tlog=Server.CreateObject(nasterloglognaster);
			mo_user=Session("EX_UTILISATEUR");
			if(mo_user!="")
			{
				tlog("F","FIN SESSION ******",true);	
				Session("EX_UTILISATEUR")="";
				mo_user="";
			}
			mo_tlog=r;
		}
		else			//  OFF
		{
			tlog("O","/TLOG OFF******/",true);	
			mo_tlog=null;
			Session("MO_ISLOG")=false;
		}
		break;
//---------------------------TLOGBEG name : timing-----------
	  case "TLOGBEG":   
		c=getword(mo_ctx);
		Session(c)=(new Date()).getTime();
		break;
		
//---------------------------TLOGEND name : timing-----------
	  case "TLOGEND":   
		c=getword(mo_ctx);
		t=Session(c);
		if(t!=null)
		{
			tlogtimer(c,(new Date()).getTime()-t);
			Session(c)=null;
		}
		break;
		
//---------------------------SQLTRACE <connection> TRUE/FALSE    oracle trace-----------------------------------------
	  case "SQLTRACE":
        src=getword(mo_ctx);		// name of datasource
        cnx=Session(src);
		c=getword(mo_ctx);			// TRUE  /  FALSE
		cmd=Server.CreateObject("ADODB.Command");
		with (cmd)
		{
    		ActiveConnection =cnx;
			CommandType = adCmdText;
			CommandText = "alter session set sql_trace = "+c;
			Execute();
		}
		cnx=null;
		cmd=null;
		break;
		    
//---------------------------------COMMANDES "label=option/..."  ---------------COMMANDES
	  case "COMMANDES":
	    commandes=getstring(mo_ctx);
	    lang=getword(mo_ctx);
	    if(lang==";")
	    {
			Session("MO_COMMANDES")=commandes;
			mo_ctx("ptr")--;
		}
		else if(Session("MO_LANGUE").toUpperCase()==lang)
			Session("MO_COMMANDES")=commandes;
	    break;
//---------------------------------NOCOMMANDES  --------------------------------------------NOCOMMANDES
	  case "NOCOMMANDES":
		Session("MO_COMMANDES")=null;
	    break;
//---------------------------------LOADXML <fic>.xml-----------------------------------LOADXML
	  case "LOADXML":
		killxmls();		
		objXMLS=Server.CreateObject("xmlservice.document2");  
		page=geturl(mo_ctx);
		Session("MO_NOMXML")=mo_page;
	    page= Server.MapPath(page);	    
	    try{
	    r=objXMLS.loadxml(page);
        objXMLS.Service.Action="moniteur.asp";
        Session("MO_XMLS")=objXMLS.XmlDocument;
		}catch(e){error("10 échec(c) load xml "+page + '\n\r' + e);}
		break;
//---------------------------------SWAPXML -----------------------------------SWAPXML
	  case "SWAPXML":
	    tmp=Session("MO_XMLS");
	    Session("MO_XMLS")=Session("MO_XMLS2");
	    Session("MO_XMLS2")=tmp;
	    tmp=null;
	    break;
//---------------------------------CLOSEXML ----------------------------------------- CLOSEXML
   	  case "CLOSEXML":
		killxmls();		
		break;		
//---------------------------------TRANSFORM [<fic>.xsl,language]----------------------------TRANSFORM
	  case "TRANSFORM":
	    page= geturl(mo_ctx);
	    lang=Session("MO_LANGUE");
	    
	    if(page!="")
	    {
	      page= Server.MapPath(page);
	      temp=Server.CreateObject(domDocumentProgId);
          r=temp.load(page);
          if(!r)
	        error("11 échec load xml "+page);
		  Session("MO_XSL2").stylesheet = temp;
		  temp=null;
		  oxsl=Session("MO_XSL2");
          if(getword(mo_ctx)==";")
            mo_ctx("ptr")--;
          else
            lang=getword(mo_ctx).toLowerCase();
        }
        else
          oxsl=Session("MO_XSL");
        
        plang=Server.CreateObject("xmlservice.ParameterList");
        plang.Add("lang",lang);
        
    try{
        Response.Write(XmlServiceFactory2("MO_XMLS").Service.Layout.ApplyStylesheet(oxsl,plang)); //,lang));
	}catch(e){Session("MO_ERROBJ")=e;error("12 TRANSFORM");}
        Session("MO_VIAXML")=1;
		break;
		
//---------------------------------SENDDATA <fic>.asp---------------------------------SENDDATA
	  case "SENDDATA":
	    page= Server.MapPath(geturl(mo_ctx));
   	    Response.Write(loaddata(page)) ; 
		break;
	    
//---------------------------------DISP page----------------------------------------- DISP
      case "DISP" :		// tag 
//----------------------------------RAW page------------------------------------- RAW
      case "RAW"  :    
//----------------------------------VIEW---------------------------------------- VIEW
      case "VIEW" : 
      
        mo_ctx("lastpage")=mpos;	
        mo_ctx("lastaff")=mpos;    
        copyarray(mo_ctx("laststack"),mo_ctx("stack"));

        if(verb=="VIEW")
        {
			service=getword(mo_ctx);
			sep=getword(mo_ctx);
			if(sep==";")
			{
				tpl="GEN";
				mo_ctx("ptr")--; 
			}
			else
				tpl=getword(mo_ctx).toUpperCase();      // name of template
			page=pathfile("TPL_"+tpl+".ASP");
        }  
        else
        {
			page=geturl(mo_ctx);      // name of page
        }
        
        if(verb=="DISP")
        {
			if(page!="")
			{
				MyExecute(page);
			}
		}
        else if(verb=="RAW"){		// RAW
			header();
			MyExecute(page);
			footer();
			Session("F_InControl") = false;
	 		Session("F_InPage") = false;
        }
        else{					// VIEW
			Session("SERVICE")=service;
			MyExecute(page);
			Session("F_InControl") = false;
			Session("F_InPage") = false;
        }
                   
        break;
//----------------------------------SEND [nosave][clear][noreturn][nocontent]--------------------------------------- SEND        
      case "SEND" :		 // envoi de la page
		Session("MO_NOSAVE")=option(mo_ctx,"nosave"); // don'save the form
		Session("MO_CLEAR")=option(mo_ctx,"clear"); // clear xmls
        Session("F_InControl") = false;
        Session("F_InPage") = false;
        savexmls();

		if(Session("MO_MESSEXT") && Session("MO_DROITS")("BEXT")==1)
		{
			wrapper="ASTEXT001";
			name="WR_"+wrapper;
			wrp=XSession(name);
			if(wrp==null)
				wrp=Server.CreateObject (mo_wrname+"."+wrapper); 
			cnx=Session("MO_CONN");
			if(cnx==null)
			{
				cnx= Server.CreateObject("ADODB.Connection");				
				cnx.Open (Session("MO_DATASOURCES")("MO_CONN")("constr"));
			}
			wrp.Connection=cnx;
			if(wrp.hasmsggestionnaire(Session("MO_USER")))
			{
				Response.Write("<script>alert('il y a des messages EXTRANET');</script>");
				Session("MO_MESSEXT")=false;
			}
			wrp=null;		
			cnx=null;
		}
//------------------------------ destruction of wrappers ----------------------   
		killwrapper();   
		if(!mo_persist_connection)
			closeconnections()			
		tlog("S","SEND");
		Session("MO_TIME")=mo_time;	
%>		
		<script>
			var sMO_TRANSACTION='<%=Session("MO_TRANSACTION")%>';
			var sMO_NOMPAGE='<%=rtrim(Session("MO_NOMPAGE"))%>';
			var sMO_TITLE=unescape('<%=escape(Session("MO_TITLE"))%>');
			var sMO_NBINST="<%=sMO_NBINST%>";
			var sMO_INSTANCES="<%=sMO_INSTANCES%>";
			var sMO_STACK="<%=sMO_STACK%>";
			var sMO_NBINST2="<%=sMO_NBINST2%>";
		</script>
<%		
		if(option(mo_ctx,"noreturn"))
		{
			if(save_callstack<Session("MO_CALLSTACK").length)
			{
				setcount(Session("MO_CALLSTACK"),save_callstack+1);
			    Session("MO_CTX")=pop(Session("MO_CALLSTACK"));
				Session("MO_TRANSACTION") =Session("MO_CTX")("name");
				Session("MO_CALLER") =Session("MO_CTX")("caller");
			}
			if(save_savestack<Session("MO_SAVESTACK").length)
			{
				setcount(Session("MO_SAVESTACK"),save_savestack+1);
				restore();
			}
		}
		if(option(mo_ctx,"nocontent"))
		{
		    Response.Status="204 NO CONTENTS";	//does not erase last page
		    ResponseEnd();
		}

		if (Application("message")!=null && Application("message")!="")
		{
			if(Session("MO_MESSLU")==null || !Session("MO_MESSLU"))
			{
				Response.Write("<script>alert('"+Application("message")+"');</script>");
				Session("MO_MESSLU")=true;
			}
		}
		else
			Session("MO_MESSLU")=false;
		
		if(Session("MO_VIAXML")==1 && !Session("MO_NOSAVE"))
			Session("MO_VIAXML")=2;
		else
			Session("MO_VIAXML")=0;
        ResponseEnd();
//-----------------------------------EXEC page--------------------------------------- EXEC
      case "EXECP" :
      case "EXEC" :
        if(Session("MO_INTRANS") && Session("MO_VBSERR"))  // on a eu un onerror
        {
			geturl(mo_ctx);  // skip next word
			break;
		}

        mo_ctx("lastpage")=mpos;  
        page=geturl(mo_ctx);
		tlog("X",mo_page);       
        MyExecute(page);		
		tlog();  
		if(verb=="EXECP")
			Session("MO_NOMPAGE")=mo_page+" ";
		else if(Session("MO_NOMPAGE").indexOf(" ")==-1 && mo_page!="PGL_COMMAND.ASP")
			Session("MO_NOMPAGE")=mo_page;
		
        break;
                
//----------------------------IMBRICATE trans-------------------------------------------- IMBRICATE
      case "IMBRICATE":  
//-----------------------------CALL trans---------------------------------------------CALL
      case "CALL" :     

		mo_ctx("issaved")=true;
        xpush(mo_ctx,verb,mo_ctx("lastpage"));  
//------------------------------SAVE-------------------------------------------- SAVE
      case "SAVE" :     
		save();		
        break;
        
//----------------------------RESTORE---------------------------------------------- RESTORE
      case "RESTORE" :  
        restore();
        break;
        
//-----------------------------TRANS transaction--------------------------------------------- TRANS
      case "TRANS" :  
//-----------------------------TRANS transaction--------------------------------------------- TRANS
      case "BRANCH" :  // call and empty stack
//-----------------------------TRANSFER transaction------------------------------------------ TRANSFER
      case "TRANSFER" :  // call and return to preceding caller   
        t=gettrans(mo_ctx);   
		tlog("T",t,true);
        p=texttrans(t);  
        if(p==null||p=="")
			error("13 Transaction "+t+" inexistante");
        nctx=newArray2();   
        if(verb=="TRANS")
			inittrans(nctx,p,t,Session("MO_TRANSACTION"));  
		else
			inittrans(nctx,p,t,""); 
			
        if(verb=="TRANS")
			push(Session("MO_CALLSTACK"),mo_ctx);  
        else if(verb=="TRANSFER")
        {
			nctx("transfer")=true;
			nctx("request")=mo_ctx("request");
		}
        Session("MO_CTX")=nctx;  
        if(verb=="BRANCH")
        {
			Session("MO_SAVESTACK")=newArray();
			Session("MO_CALLSTACK")=newArray();
		}
        play(nctx);   
        if(verb=="BRANCH")
        {
			Session.Abandon();
			Response.Clear();
			ResponseEnd();	
		}		
	    if(nctx("request") && nctx("transfer"))
			pop(Session("MO_CALLSTACK"));
	    if(Session("MO_CALLSTACK").length>0)
			mo_ctx=pop(Session("MO_CALLSTACK")); 
	    if(mo_ctx("issaved"))
	    {
			restore();
			mo_ctx("issaved")=false;
        }       
		Session("MO_TRANSACTION") =mo_ctx("name");
		Session("MO_CALLER") =mo_ctx("caller");
        Session("MO_CTX")=mo_ctx;    
	    mo_skipcase=false;
	    continue;
        
//-------------------------------CASE condition------------------------------------------- CASE
      case "CASE" :    // test
      
        if(mo_skipcase ){ // if already got a CASE, then skip
          skip(mo_ctx);   
          continue;
        }
      
        not=option(mo_ctx,"NOT");  // CASE NOT
        c=getcond(mo_ctx);		  // read condition
        if(c=="EMPTY")
			c="";
        d=Session("MO_DROITS");   // or DR:<user right code>
        if (c.substr(0,3)=="DR:"){
          if(d(c.substr(3))!=1^not){ // dictionary value is 1 when permitted
            skip(mo_ctx);			
            continue;         
          }
        }
        else					// return code
          if((c!=mo_rc)^not){
            skip(mo_ctx);		// fail
            continue;
          }
        xpush(mo_ctx,verb,""); // stack "CASE"
        break;
        
 //----------------------------OTHER---------------------------------------------- OTHER
     case "OTHER" :								
        if(mo_skipcase ){
          skip(mo_ctx);
          continue;
        } 
        xpush(mo_ctx,verb,""); 
        break;
          
//-----------------------------REPEAT instr--------------------------------------------- REPEAT
      case "REPEAT" :		// loop (stopped by JUMP or BREAK)
      
        xpush(mo_ctx,verb,mpos); // save REPEAT start
        break;
        
        
//-------------------------------ON cond instr------------------------------------------- ON
      case "ON" :          // async event
      
        c=getcond(mo_ctx);   
        p=skip(mo_ctx);    
        mo_ctx("ons")(c)=p;
        break; 

//-------------------------------CONTINUE------------------------------------------- CONTINUE
      case "CONTINUE" :        
        break;
//-------------------------------LABEL label------------------------------------------- LABEL
      case "LABEL" :		
      
        skip(mo_ctx);    
        break; 
        
        
//-------------------------------EVAL script ---- §  §
      case "§":     // evaluate javascript
      
        mo_ctx("ptr")=mpos;  
        txt=geteval(mo_ctx);   // read  script
        try{
          eval(txt);
        }  
        catch(e){
          error("14 erreur dans l'exécution du script: \n"+txt);
        }
	    mo_skipcase=false;
	    continue;
//-------------------------------EVAL script : result into MO_RETURN----- EVAL
      case "EVAL" :
        txt=geteval(mo_ctx);   // read script
        try{
          x=eval(txt);
        }  
        catch(e){
          error("15 erreur dans l'exécution du script: \n"+txt);
        }
        if(x!=null)
        {
			try {Session("MO_RETURN")=x.toString();} 
			catch(e){Session("MO_RETURN")="";}
		}
        break;
        
//---------------------------------JUMP label----------------------------------------- JUMP
      case "JUMP" :        
	    mo_more=dojump(mo_ctx,"")    
        break;
       
//---------------------------------JUMPPREV [label]----------------------------------------- JUMPPREV
      case "JUMPPREV" : // jump to preceding label or any label
        mo_more=dojumpprev(mo_ctx,"") 
        break;        
      
//----------------------------------JUMPNEXT [label]---------------------------------------- JUMPNEXT
       case "JUMPNEXT" :   // same forward
        mo_more=dojumpnext(mo_ctx,"")
        break;
     
//----------------------------------READTRANS fic---------------------------------------- READTRANS
      case "READTRANS" :  // read transaction file     
        file=Server.MapPath(geturl(mo_ctx));   // file name
        readtrans(file);
        break;


//---------------------------------------------------------------------------------------------
//---------------------------------------------- CONNECTIONS and TRANSACTIONS
//---------------------------------------------------------------------------------------------
        
//---------------------------------DEFINEDATASOURCE name="string"-------------------------DEFINEDATASOURCE
      case "DEFINEDATASOURCE" :
        src=getword(mo_ctx);		// datasource name
        getword(mo_ctx);		// skip "="
        str=getstring(mo_ctx);		// connection string
        txt=Session("MO_SCHEMA");
        if(src=="MO_CONN" && txt!="")
        {
			str=str.replace(/User Id=\w*;/i,"User Id="+txt+";");
			txt=Session("MO_PASSWORD");
			str=str.replace(/Password=\w*;/i,"Password="+txt+";");
			txt=Session("MO_INSTANCE");
			str=str.replace(/Data Source=\w*/i,"Data Source="+txt);
		}
        
        a=newArray2();
        a("constr")=str;
        a("nbusers")=0          // usage count
        a("isintran")=false;	// pending transaction
        a("translevel")=0;		//transaction level
        a("transconn")=newArray() ;        // transaction connections
        
        Session("MO_DATASOURCES")(src)=a;  
        break;
//---------------------------------GETCONNECTION name [,wrapper,..]----------------------------------------GETCONNECTION 
      case "GETCONNECTION2" :		// 2nd connection
         bcon2=true;
      case "GETCONNECTION" :
        src=getword(mo_ctx);		// datasource name
        // look for avaiable connection
        a=Session("MO_DATASOURCES")(src);
        if(a==null)
          error("16 no connection");
        
		if(a("nbusers")>0)
		{          
          cnx=Session(src);
          if(!mo_persist_connection)
			 a("nbusers")++;	
        }
        else
        {
		  a("nbusers")=1;
		  cnx= Server.CreateObject("ADODB.Connection");
          cnx.Open (a("constr"));
          Session(src)=cnx;
        }
        
        arr=getlist(mo_ctx,false);  // arguments (wrappers)
        for(wrapper in arr)
        {
			name="WR_"+wrapper;
			if(XSession(name)==null)
			{
				XSession(name)=Server.CreateObject (mo_wrname+"."+wrapper);
			}
			Session("MO_DICWR")(name)=true;			
			XSession(name).Connection = cnx;
        }  
        cnx=null;
        break;
//---------------------------------GLOINIT name ----------------------------------------GETCONNECTION 
      case "GLOINIT" :
        src=getword(mo_ctx);		// datasource name
        cnx=Session(src);
        cnx=null;
		break;
//---------------------------------CLOSECONNECTION name-----------------------------------------CLOSECONNECTION 
      case "CLOSECONNECTION2" :		
        bcon2=true;
      case "CLOSECONNECTION" :
		hasconnerror(false);
        src=getword(mo_ctx);		// datasource name

        if(!mo_persist_connection || bcon2)
        {
			a=Session("MO_DATASOURCES")(src);
			if(--a("nbusers")!=0)          
			  break;
			Session(src).Close();
			Session(src)=null;
			a("isintran")=false;
			Session("MO_INTRANS")=isintrans();        
		}
	    mo_skipcase=false;
	    continue;
//-----------------------------BEGINTRANS [datasource,..]---------------------------------------BEGINTRANS
      case "BEGINTRANS" :    // database transaction 
        arr=getlist(mo_ctx,true);  //  arguments (connections)
        ds=Session("MO_DATASOURCES");
        for(key in arr)
        {
          a=ds(key);   
          if(a("nbusers")>0 && a("translevel")++==0)
          {
            Session(key).BeginTrans();
			xxxtran++;            
            a("isintran")=true;	
            a("transconn")=arrlisttovec(arr) ;  
          }
        }  
        Session("MO_INTRANS")=true;
		Session("MO_VBSERR")=false;
        break;
        
        
//----------modifié-------------COMMIT [datasource,..]-------------------------------------------- COMMIT
      case "COMMIT" :
		if(! hasconnerror(true))
		{
	        arr=getlist(mo_ctx,true);  // arguments (connections)
	        ds=Session("MO_DATASOURCES");
	        
	        for(key in arr)
	        {
	          a=ds(key);   
	          if(a("nbusers")>0 &&(a("translevel")==0 || --a("translevel")==0)) 
	          {
	            isin=true;
	            ltr=a("transconn");
	            cnt=ltr.length;
	            for(x=0;x<cnt;x++)
	              if(arr[ltr(x)]==null)
	              {
	                isin=false;
	                break;
	              }
	            
	            
	            if(a("isintran") && isin) // if no other connections
	            {
				  try{ Session(key).CommitTrans();		xxxtran--; } 
				      catch(e)
				      {
				      null;
				      }
	              a("isintran")=false;	
	            }
	          }
	        }  
	        Session("MO_INTRANS")=isintrans();
			Session("MO_VBSERR")=false;
	        break;
	    }       
		Session("MO_VBSERR")=false;
//-------------------------------ROLLBACK [datasource,..]------------------------------------------ ROOLBACK
      case "ROLLBACK" :
        arr=getlist(mo_ctx,true);  // arguments (connections)
        ds=Session("MO_DATASOURCES");
        for(key in arr)
        {
          a=ds(key);   
          if(a("nbusers")>0 )
          {
            if(a("isintran"))
            {
			  try{ Session(key).RollbackTrans(); xxxtran--;} catch(e){}
              a("isintran")=false;	
              a("translevel")=0;
            }
          }
         // rollback all connections 
	      ltr=a("transconn");
	      cnt=ltr.length;
	      for(x=0;x<cnt;x++)
          {
            key2=ltr(x)
            a2=ds(key2);   
            if(a2("nbusers")>0 )
            {
              if(a2("isintran"))
              {
 			    try{ Session(key2).RollbackTrans(); xxxtran--; } catch(e){}
                a2("isintran")=false;	
                a2("translevel")=0;
              }
            }
          }  
        }  
        Session("MO_INTRANS")=isintrans();
        break; 
        
//---------------------------------------------------------------------------------------------
//----------------------------- WRAPPER  access to data thru generated COM component
//---------------------------------------------------------------------------------------------
        
//---------------------------- GETWRAPPER wrapper1,..---------------------------------------GETWRAPPER
      case "GETWRAPPER" :    // instanciate plwrapper.<wrapper1>
        arr=getlist(mo_ctx,false);  //  arguments (wrappers)
        for(wrapper in arr)
        {
			name="WR_"+wrapper;
			if(XSession(name)==null)
			{
				XSession(name)=Server.CreateObject (mo_wrname+"."+wrapper);
			}
			Session("MO_DICWR")(name)=true;
        }  
        break;
//---------------------------- CLOSEWRAPPER wrapper1,..---------------------------------------CLOSEWRAPPER
      case "CLOSEWRAPPER" :    // destruction plwrapper.<wrapper1>
        arr=getlist(mo_ctx,false);  //  arguments (wrappers)
        for(wrapper in arr)
        {
			name="WR_"+wrapper;
			XSession(name)=null;
			if(Session("MO_DICWR").Exists(name))
				Session("MO_DICWR").Remove(name);
        }  
	    mo_skipcase=false;
	    continue;
//---------------------------- UPDATE ---------------------------------------
      case "UPDATE" :    // call standard  Modify method
        wrapp=XSession("WR_"+getword(mo_ctx));		// wrapper name
        ors=wrapp.RecordStructure;
        try{
		XmlServiceFactory2("MO_XMLS").Service.ElementsToRecordset(ors);
		}catch(e){Session("MO_ERROBJ")=e;error("17 update");}		
		try{
		wrapp.Modifie(ors);
		}catch(e){}
        break;                  
//---------------------------- DELETE ---------------------------------------
      case "DELETE" :    // call standard  Delete method
        wrapp=XSession("WR_"+getword(mo_ctx));		// wrapper name
        ors=wrapp.RecordStructure;
        try{
		XmlServiceFactory2("MO_XMLS").Service.ElementsToRecordset(ors);
		}catch(e){Session("MO_ERROBJ")=e;error("18 delete");}		
		try{
		wrapp.Delete(ors);
		}catch(e){}
        break;                  
 //---------------------------- INSERT ---------------------------------------
      case "INSERT" :    // call standard  Insert method
        wrapp=XSession("WR_"+getword(mo_ctx));		// wrapper name
        ors=wrapp.RecordStructure;
        try{
		XmlServiceFactory2("MO_XMLS").Service.ElementsToRecordset(ors);
		}catch(e){Session("MO_ERROBJ")=e;error("19 insert");}		
		try{
		wrapp.Cree(ors);
		}catch(e){}
		try{
		XmlServiceFactory2("MO_XMLS").Service.RecordsetToElements(ors);
		}catch(e){Session("MO_ERROBJ")=e;error("20 insert");}		
        break;                  
        
//---------------------------------------------------------------------------------------------
//---------------------------------------------- RETURN CODES processed as INSTRUCTIONS
//---------------------------------------------------------------------------------------------
      case "ERROR":
        if(mo_ctx("lastaff")<0)
        {
          Session("MO_RETURN")="ERROR";
          return;		
        }
        mo_ctx("ptr")=mo_ctx("lastaff");
        mo_ctx("stack")=clone(mo_ctx("laststack"));
        break;

      case "AGAIN" :
        if(mo_ctx("lastpage")<0)
        {
          Session("MO_RETURN")="AGAIN";
          return;		
        }
        mo_ctx("ptr")=mo_ctx("lastpage");
        break;          
      
      case "LOOP":
      case "BREAK":
        bfound=false;      
        while((ct=xpop(mo_ctx))!=null ){
          if(ct("verb")=="REPEAT"){
            bfound=true;
            mo_ctx("ptr")=ct("data");
            if(verb=="BREAK")
              skip(mo_ctx);
            break;
          }  
        }
        if(!bfound)
        {
          Session("MO_RETURN")=verb;
          return;
        }
        break;

      case "ABORT" :
      case "RETURN" :
        mo_more=false;
        if(mo_ctx("name")=="request")
          Session("MO_RETURN")="ABORT";		
        continue;
        
      case "ABANDON" :  // stop session
		doAbandonAppli();
		
	 case "ENDMODAL" :
	    if(Session("MO_MODAL")!="NO")
	    {
			if(Session("MO_CALLSTACK").length>0)
			{
				Session("MO_CTX")=pop(Session("MO_CALLSTACK")); 
				if(Session("MO_CTX")("issaved"))
				{
					restore();
				    Session("MO_CTX")("issaved")=false;
				}
				Session("MO_TRANSACTION") =Session("MO_CTX")("name");
				Session("MO_CALLER") =Session("MO_CTX")("caller");
			}		    
			Session("MO_MODAL")="NO";
%>
	<script> 
		document.cookie = "modal=NO;path="+escape("/");
		document.cookie = "lastnum=0;path="+escape("/");
		document.cookie = "result=<%=escape(Session('MO_RESULT'))%>;path="+escape("/");
		window.close();
	</script>
<%		    Session('MO_NPAGEOUT')=null;
			doEnd();
		}
		continue;

      case "LANGUAGE" :  // language
        Session("MO_LANGUE")=getword(mo_ctx).toLowerCase();		// code language
		break;
        
      default : // instruction error
      
        error("21  script error due to unknown verb: "+verb);
        
    }			
    
    //-----------------------------------------------------------------------
    
    mo_skipcase=false;
    mo_rc="";
    if((r=Session("MO_REQUEST"))!="")			// return from Server.exec
    {
      nctx=newArray2();   
      inittrans(nctx,r,"request");  
	  nctx("request")=true;   
      push(Session("MO_CALLSTACK"),mo_ctx);  
      Session("MO_CTX")=nctx;  
      Session("MO_REQUEST")="";
      play(nctx);   
	  if(nctx("request") && nctx("transfer"))
	     pop(Session("MO_CALLSTACK"));
      mo_ctx=pop(Session("MO_CALLSTACK")); 
      Session("MO_CTX")=mo_ctx;    
    }
  } 
  //------------------------------   END LOOP----------------------------
}
//			END PLAY

//----------------------------------------------------------------------------
//----------------------------------------------------------------------------
//-------------------HELPER FUNCTIONS ----------------------------------------
//----------------------------------------------------------------------------
//----------------------------------------------------------------------------
//---------------------------------------------------------------------------- getword
function getword(ctx,noeval){
//----------------------------------------------------------------------------
//  get token
  var ret,s,i,cont,inword,c,rpos;
  ret="";			
  s=ctx("prg");
  i=ctx("ptr");
  rpos=i;			
  cont=true;		
  inword=false;
  while(cont){
    c=s.charAt(i++);
    switch (c){
      case "" :		
        cont=false;
        break;
      case " ":		
      case "\t":
      case "\n":
      case "\r":
        if (inword)
          cont=false;  
        break;
      case ";" :		
      case "(" :
      case ")" :
      case "=" :
      case "*" :
      case "§" :
      case "," :
        if (inword)
          i--;		
        else
          ret=c;    
        cont=false; 
        break;
      default:		
        ret+=c;
        inword=true;
        break;
    }
  }
  ctx("ptr")=i;	
  if(noeval==null && ret=="§"){  
    ctx("ptr")=rpos;		
    txt=geteval(ctx);   
    try{
      ret=eval(txt).toString();  
    }  
    catch(e){}
  }  
  
    ret=ret.toUpperCase();  
  return ret;
 }
//---------------------------------------------------------------------------- getlist
function getlist(ctx,bds){
//----------------------------------------------------------------------------
//  get list  (connection if bds=true),.... 
var src,ds,key,keys,arr,i;

arr=new Array();
src=getword(ctx);
if(src==";")
{
	ctx("ptr")--;   
	if(bds)
	{
	  ds=Session("MO_DATASOURCES"  );  
	  keys=ds.Keys().toArray();
	  for (i in keys)
	    arr[keys[i]]=0;
	 }
	 return arr;
}
while (true)
{
  if(src!="," && src!=";" )
  {
	arr[src]=0;
	src=getword(ctx);
  }
  if(src==";")
  {
    ctx("ptr")--;   
    return arr;
  }
  src=getword(ctx);
}        
}//---------------------------------------------------------------------------- isintrans
function isintrans(){
//----------------------------------------------------------------------------
var ds,keys,i;

ds=Session("MO_DATASOURCES"  );  
keys=ds.Keys().toArray();
for (i in keys)
  if(ds(keys[i]).isintran)
    return true;
return false;
}
//---------------------------------------------------------------------------- getstring
function getstring(ctx){
//----------------------------------------------------------------------------
//  get string "................" or eval § ... §
  var ret,s,i,cont,inword,c,del;
  ret="";			
  del="";
  s=ctx("prg");
  i=ctx("ptr");
  cont=true;		
  while(cont){
    c=s.charAt(i++);
    switch (c){
      case "" :		
        cont=false;
        break;
      case '"':
      case '§':
        if (del==c)
        {
          cont=false;  
          break;
        }
        else
			if (del=="")
			{
	          del=c;
	          break;
	        }
	        else
				if(del=='"')		
				{
					del="";
					break;
				}
      default:		
        if(c=="\t")	
			c=" ";  
        if(del!="")
          ret+=c;
        break;
    }
  }
  ctx("ptr")=i;	
  if(del=="§")
    try{
      ret=eval(ret).toString();  
    }  
    catch(e){}
  return ret;
 }
  

//---------------------------------------------------------------------------- geteval
function geteval(ctx){
//----------------------------------------------------------------------------
// read § javascript ... §

  var s,i1,i2;
  
  s=ctx("prg");
  i1=s.indexOf("§",ctx("ptr"))+1;   
  if((i2=s.indexOf("§",i1))<0) 
    scripterror();			
  ctx("ptr")=i2+1;                 
  return s.substring(i1,i2);
 }

//---------------------- 2nd level functions ------------------
function getverb(ctx){return getword(ctx,true);} 
function gettrans(ctx){return getword(ctx);}          
function getcond(ctx){return getword(ctx);} 
//------------------------------------------------------- geturl         
function geturl(ctx)
//-------------------------------------------------------         
{
  var url,idx,name;
  
  url= getword(ctx);    // token
  if(url==";")
  { ctx("ptr")--; 
    return "";
  }
  if(url.indexOf(".")<0)
    url+=".ASP";
  mo_page=url;
  if((name=pathfile2(url))==null)  
    return url;             
  return name;   
}          
        
//------------------------------------------------------------------------ option
function option(ctx,word){
//----------------------------------------------------------------------------
// get optional  token , when not found, restore pointer and returns false
  var i,w;
  i=ctx("ptr");
  w=getword(ctx);
  if(word.toUpperCase()==w)
    return true;
  ctx("ptr")=i;
  return false;
}

//----------------------------------------------------------------------- skip
function skip(ctx){
//----------------------------------------------------------------------------
// skip until ";" 
  var s,i,nbpar,ineval,cont,c,idep;

  s=ctx("prg");
  i=ctx("ptr");
  idep=i;
  cont=true;
  nbpar=0;
  
  while(cont){
    c=s.charAt(i++); 
    switch(c){
      case ";" :
        if(nbpar==0 )
          cont=false;
        break;
      case "(" :
        nbpar++;
        break;
      case ")" :
        nbpar--;
        break;
      case "§" :
        ctx("ptr")=i-1;
        geteval(ctx);   // skip §.....§
        i=ctx("ptr");
        break;
      case '"' :
        ctx("ptr")=i-1;
        getstring(ctx);   // skip §.....§
        i=ctx("ptr");
        break;
    }          
  }
  ctx("ptr")=i-1;
  return s.substring(idep,i);
}
//----------------------------------------------------------------------- skipcom
function skipcom(ctx){
//----------------------------------------------------------------------------
// skip comment until cr/lf  
  var s,i,c,idep,lf;

  s=ctx("prg");
  i=ctx("ptr");
  idep=i;
  lf=String.fromCharCode(10);
  
  while((c=s.charAt(i++))!=lf && c!="");
  ctx("ptr")=i;
  return s.substring(idep,i);
}
 
//--------------------------------------------------------------------- xpush
function xpush(ctx,verb,data){
//----------------------------------------------------------------------------
//  push verb and data
  var ct;
  ct=newArray2();
  ct("verb")=verb;
  ct("data")=data;
  push(ctx("stack"),ct);
}  

//---------------------------------------------------------------------- xpop
function xpop(ctx){    
//----------------------------------------------------------------------------
  return pop(ctx("stack"));
}  

//----------------------------------------------------------------------- clone
function clone(arr){
//----------------------------------------------------------------------------
  var ret,i;
  if(mo_caprockvect)
  {
	ret=newArray();
	for(i=0;i<arr.length;i++)
	  ret(i)=arr(i);
  }
  else
  {
	ret=new Array();
	for(i=0;i<arr.length;i++)
	  ret[i]=arr[i];
  }
  return ret;
}    

//----------------------------------------------------------------------- copyarray
function copyarray(adst,asrc)
//----------------------------------------------------------------------------
{
  var l=asrc.length,i;
  if(mo_caprockvect)
  {
	for(i=0;i<asrc.length;i++)
	  adst(i)=asrc(i);
  }
  else
  {
	for(i=0;i<asrc.length;i++)
	  adst[i]=asrc[i];
  }
  setcount(adst,l);  
}

//------------------------------------------------------------------------ save
function save()
{
	var arr=newArray2();
	for(i=1;i<=Session.Contents.Count;i++){
		k=Session.Contents.Key(i)
		s=k.substr(0,3);             
		s=s.toUpperCase();
		if(s!="MO_" && s!="GL_" && s!="MEM_GL_" && s!="WR_") // variables with names like GL_ or WR_ are global and not saved
				arr(k)=Session.Contents.Item(i);
	}
	savexmls();
	arr("MO_XMLS")=Session("MO_XMLS"); 
	arr("MO_XMLS2")=Session("MO_XMLS2"); 
    
	arr("MO_COMMANDES")=Session("MO_COMMANDES");
	arr("MO_TITLE")=Session("MO_TITLE");
	if(Session("MO_OLD_MENU")!="")
	{
		arr("MO_MENU")=Session("MO_OLD_MENU");
		Session("MO_OLD_MENU")="";
	}
	arr("MO_LASTRETURN")=Session("MO_LASTRETURN");
	arr("MO_NOSAVE")=Session("MO_NOSAVE");
	push(Session("MO_SAVESTACK"),arr); 
	killxmls();		
}
//------------------------------------------------------------------------ restore
function restore(){   
//----------------------------------------------------------------------------
  var arr,i,keys,key;
  arr=pop(Session("MO_SAVESTACK"));
  keys=arr.Keys().toArray();
  for(i in keys)
  {
    key=keys[i];
    Session(key)=arr(key);
  }
  restorexmls();    
}        

//----------------------------------------------------------------------getlabels
function getlabels(ctx){
//----------------------------------------------------------------------------
  var s,i,lab,rpos,arr;
  rpos=ctx("ptr");
  ctx("ptr")=0;     
  s=ctx("prg").toUpperCase();
  while((i=s.indexOf("LABEL",ctx("ptr")))>=0){
    ctx("ptr")=i+5;
    lab=getword(ctx);    
    if(lab==";")
      lab="NONAME";      
    else
      getword(ctx);   
    arr=newArray();
    if(mo_caprockvect)
    {    
        arr(0)=lab;
        arr(1)=ctx("ptr");
    }
    else
    {
        arr[0]=lab;
        arr[1]=ctx("ptr");
    }
    push(ctx("labels"),arr); 
  }
  ctx("ptr")=rpos;
}

//-------------------------------------------------------------------dojump
function dojump(ctx,lab){
//----------------------------------------------------------------------------
  var lab,i;
    if (lab=="")
	  lab=getword(ctx);
	for(i=0;i<ctx("labels").length;i++)
	  if(ctx("labels")(i)(0)==lab)
	  {
	    jump(ctx,ctx("labels")(i)(1));   
	    return true;
	  }					
	mo_jump=0;           
	mo_label=lab;        
    return false;
}	

//------------------------------------------------------------------dojumpprev
function dojumpprev(ctx,lab){
//----------------------------------------------------------------------------
  var lab,i;
    if (lab=="")
      lab=getword(ctx);
    if(lab==";")  
      ctx("ptr")--;
    pos=0;
    for(i=0;i<ctx("labels").length;i++)
    {
      p=ctx("labels")(i)(1); 
      if(p>ctx("ptr"))
        break;  
      if(lab==";" || ctx("labels")(i)(0)==lab)
        pos=p;		
    }
    if(pos==0){
	  mo_jump=-1;   
	  mo_label=lab;
      return false;
    }
    jump(ctx,pos);		
    return true;
}    

//------------------------------------------------------------------ dojumpnext 
function dojumpnext(ctx,lab){
//----------------------------------------------------------------------------
    var lab,i;
    if (lab=="")
      lab=getword(ctx);
    if(lab==";")
      ctx("ptr")--;
    pos=0;
    for(i=0;i<ctx("labels").length;i++)
    {
      p=ctx("labels")(i)(1);
      if(p<ctx("ptr"))    
        continue;        
      if(lab==";" || ctx("labels")(i)(0)==lab){
        pos=p;
        break;    
      }
    }
    if(pos==0){
	  mo_jump=1;
	  mo_label=lab;
      return false;
    }
    jump(ctx,pos);
    return true;
}    
 

//--------------------------------------------------------------- inittrans
function inittrans(nctx,p,name,caller){
//----------------------------------------------------------------------------
// initialize script
  var oldctx=Session("MO_CTX"),key;
  nctx("name")=name;			// name
  nctx("prg")=p;				// text
  nctx("ptr")=0;				//  instruction pointer
  nctx("stack")=newArray(); 	// stack instructions
  nctx("laststack")=newArray();
  nctx("ons")=newArray2();		// conditions ON
  nctx("labels")=newArray();  // labels
  nctx("lastpage")=-1;			// last executed page
  nctx("lastaff")=-1;			// last displayed page
  nctx("caller")=caller;		// calling transaction 
  nctx("issaved")=false;		// flag context is saved
  nctx("request")=false;		
  nctx("transfer")=false;		// flag transfer 
  if (oldctx!=null)
  {
    var arr=oldctx("ons").Keys().toArray();
	for(var c in arr)
	{
		key=arr[c];
		nctx("ons")(key)=oldctx("ons")(key);
	}
  }
}

//--------------------------------------------------------------- loaddata
function loaddata(fic){
//----------------------------------------------------------------------------
// load XML datas into XML,HTML,ASP
	var w,xmldoc,xmldoc2,List,i,r,xml;

	xmldoc=XmlServiceFactory2("MO_XMLS");
	xmldoc2=Server.CreateObject(domDocumentProgId);
	r=xmldoc2.load(fic);
	if(!r)
	  error("22 échec load xml "+fic);

	List = xmldoc2.selectNodes(".//*[@name]");
	for (i=0;i<List.length;i++)
	{
	  xmlnode2 =List[i];
	  name=xmlnode2.getAttribute("name");
	  try{
        xmlnode= xmldoc.Service.ElementList.Item(name);
		xmlnode2.setAttribute( "value",xmlnode.DataList.Item(1).Value);		 
	  }catch(e){}
	}
	xml= xmldoc2.xml;
	xmldoc=null;
	xmldoc2=null;
	List=null;
	
	return xml;
}
//------------------------------------------------------------------- jump
function jump(ctx,npos){
//----------------------------------------------------------------------------
  var pos,i,npar=0,s,ct,ineval;
  ineval=false;
  s=ctx("prg");
  pos=ctx("ptr");
  
  flushstack(ctx);  
  
  if(npos>pos){					
    for(i=pos;i<=npos;i++){
      c=s.substr(i,1);
      if(c=="§"){
        ctx("ptr")=i;
        geteval(ctx);
        i=ctx("ptr");
        continue;
      }
      if(c==")"){
        if(npar==0){
          xpop(ctx);   
		  flushstack(ctx);
        }
        else
          npar--;		
      }
      else if(c=="(")
        npar++;
    }
    while(npar--)			
      xpush(ctx,"BLOCK","");  
  }

  else{							
    for(i=pos;i>=npos;i--){
      c=s.substr(i,1);
      if(c=="§"){
		ineval= !ineval;
      }
      if(ineval)
        continue;
      if(c=="("){
        if(npar==0){
          xpop(ctx);   
          flushstack(ctx);
        }
        else
          npar--;
      }
      else if(c==")")
        npar++;
    }
    while(npar--)
      xpush(ctx,"BLOCK","");
  }

  ctx("ptr")=npos;
}  


//------------------------------------------------------------------------------flushstack
function flushstack(ctx){
//------------------------------------------------------------------------------
//    pop until BLOCK (parenthesis) then push
  var ct;
  while((ct=xpop(ctx))!=null) 
    if(ct("verb")=="BLOCK"){
      xpush(ctx,"BLOCK","");
      break;
    }
}
//-------------------------------------------------------------------------- READTRANS
function readtrans(file)  // read transaction file
//-------------------------------------------------------------------------- 
{
	var fs,strm,txt,temp,i
	if(!Session("MO_SESSTRANS"))    // lock if transactions dictionary in application oject
	      Application.Lock ();
	fs = Server.CreateObject("Scripting.FileSystemObject"); // file system
	strm = fs.OpenTextFile(file,1);  // stream
	txt=strm.ReadAll();    // text
	strm.Close();
	temp=newArray2();
	temp("ptr")=0;
	temp("prg")=txt;
	txt=txt.toUpperCase();   
	while((i=txt.indexOf("DEFTRANS",temp("ptr")))>=0)  
	{
		  temp("ptr")=i+8;
		  name=getword(temp);     
		  getword(temp);          // ";"
		  i=txt.indexOf("ENDDEF",temp("ptr"));   
		  if(i<0)
				break;   
		  TRANS(name)=temp("prg").substring(temp("ptr"),i);
		  i+=6;
		  temp("ptr")=i;
	}
	if(!Session("MO_SESSTRANS"))
		  Application.UnLock ();
	fs=null;
} 
   
  
//----------------------------------------------------------------------- header
 function header(){    
  MyExecute(pathfile("PGL_MENU.ASP")); }
 //----------------------------------------------------------------------- footer
function footer(){
  MyExecute(pathfile("PGL_COMMAND.ASP")); }
  
//--------------------------------------------------------------------------------doEnd 
function doEnd()          
{
	savexmls();	
    ResponseEnd();
}


//--------------------------------------------------------------------------------savexmls
function savexmls()
{
	return;
} 
//--------------------------------------------------------------------------------killxmls
function killxmls(ball)
{
	return;
}
//--------------------------------------------------------------------------------killwrapper
function killwrapper()
{
   var keys=Session("MO_DICWR").Keys().toArray();
   var name;
   for (var i=0;i<keys.length;i++)
   {
		name = keys[i];
		Session(name)=null;
		Session("MO_DICWR").Remove(name);
   }
}
function killwrapper2()
{

   var e = new Enumerator(Session.Contents);       
   var name;
   for (;!e.atEnd();e.moveNext())           //Enumerate session variables collection.
   {
		name = e.item();
		if(name.substr(0,3)=="WR_")
			Session(name)=null;
   }
}

//--------------------------------------------------------------------------------restorexmls
function restorexmls()
{
	return;
}

//-------------------------------------------------------------------------createglo
function createglo()
{
}

//---------------------------------------------------------------------------isOracle
function isOracle(cnx)
{
	return cnx.Provider.toUpperCase().substr(0,7)=="MSDAORA";
}
function closeconnections()
{
  var ds=Session("MO_DATASOURCES"),key;
  var keys=ds.Keys().toArray();
  for(var i in keys)
  {
	key=keys[i];
    if(ds(key)("nbusers")>0)
    {
	  try{ Session(key).CommitTrans();} catch(e){}
      Session(key).Close();
      ds(key)("nbusers")=0;
    }
	Session(key)=null;
  }
}
function doAbandonAppli()
{
	closeconnections();
	killwrapper();
	Response.Clear();
	Session.Abandon();
	ResponseEnd();
}

//----------------------------------------------------------------------------------------------------
//----- general functions ---------------------------------------------------------------------------
// required because there is a bug in jscript push and pop methods

function push(arr,obj){
  if(mo_caprockvect)
	arr(arr.length)=obj;
  else
	arr[arr.length]=obj;
}
function pop(arr){
  var r,i;
  i=arr.length;
  if(mo_caprockvect)
  {
	r=arr(i -1);
	if(i>0)
	{
	  arr(i -1)=null;
	  arr.removeAt(i-1);
	}
  }
  else
  {
	r=arr[i -1];
	if(i>0)
		arr.length=i-1;
  }
  return r;
}
function rtrim(s)
{
  return s.replace(/\s*$/g,"");
}


//---------------------------------------------------------------------- REQF
function REQF(str)
{
  var x=Request.Form(str);
  if((x.Item) != null)
    return x(1);
  x=Request(str);
  if((x.Item) != null)
    return x(1);
  return null;
}    
//---------------------------------------------------------------------- scripterror
function scripterror()
//---------------------------------------------------------------------- 
{
Server.Transfer(pathfile2("asp/MO_ERROR.ASP"));
}
//---------------------------------------------------------------------- hasconnerror
function hasconnerror(bretry)
//---------------------------------------------------------------------- 
{
	var lf,mess,ds,keys,key,conn,i,j,errs,err,a,bFatal,n,msg;
	lf=String.fromCharCode(10);
	mess="";
	bFatal=false;
	ds=Session("MO_DATASOURCES");  
	keys=ds.Keys().toArray();
	for (i in keys)
	{
      key=keys[i];
      a=ds(key);   
      if(a("nbusers")>0)
      {
		conn=Session(key);
		errs=conn.Errors;
		for(j=0;j<errs.Count;j++)
		{
			err=errs(j);
			msg=strip(err.Description);
			mess+=" erreur :"+err.NativeError+":"+ msg;
		}
		errs.Clear();
      }	
	}
	if(bFatal)
		error("26 "+mess);
		
	if (mess=="" && Session("MO_VBSERR"))
		mess=Session("MO_VBSERRDESC");
		
	if(mess!="")
	{
		if(!Session("MO_NOALERTERROR"))
		{
	%>
	<script>
		alert(unescape('<%=escape(mess)%>'));
	</script>
	<%
		}
		if(mess.substr(0,14)!=" erreur :20888")
			logmess("hasconnerror",mess);
		if(bretry)	
		{
			Session("MO_RETURN")="ERROR";
			Session("MO_NOALERTERROR")=false;
			Session("MO_HASERROR")=true;
			return true;
		}	
	}
	Session("MO_NOALERTERROR")=false;
	Session("MO_HASERROR")=false;	
	return false;

}
function asnum(str)
{
	str=str.replace(/ /g,"");
	str=str.replace(/,/g,".");
	return eval(str);
}

function strip(s) // takes the string between § and § remove cr
{
	var i=s.indexOf("§");
	var j=s.lastIndexOf("§");
	j=(j==-1)?s.length:j;
	s= s.substring(i+1,j);
	var re=/[\u000A\u000D]/g
	s=s.replace(re," ");
	return s;
}
function getvalue(strname)  // get a value from XMLS
{
	if(Session("MO_VIAXML")==2) 
	{
		try
		{
			return XmlServiceFactory2("MO_XMLS").Service.ElementList(strname).SingleValue;
		}
		catch(e)
		{
			return Session("RFORM")(strname.toUpperCase());

		}
	}
	else
		return Session("RFORM")(strname.toUpperCase());
}

function tlog(peventtype,pname,pimmediate)
{
	var lmilli;
	
	if(!hastolog(peventtype))
		return;
	if(mo_tlog==null)
		return;
	if(!mo_tlog)
		mo_tlog=Server.CreateObject(nasterloglognaster);

	var ldate=new Date();
	var ltime= (ldate).getTime();
	var lmoment=ldate.toLocaleString() ;
	if(mo_time!=0)
	{
		lmilli=mo_time>0?ltime - mo_time:0;
		try{
			mo_tlog.putLogData(lmoment,mo_user,lmilli,mo_eventtype,mo_name,"","",Session("MO_SERVER"));
		}catch(e){}		
	}
	if(peventtype==null)  
		mo_time=0;
	else
	{
		mo_time=ltime;
		mo_eventtype=peventtype;
		mo_name=pname;
	}
	if(pimmediate!=null)
	{
		try{
		mo_tlog.putLogData(lmoment,mo_user,0,mo_eventtype,mo_name,"","",Session("MO_SERVER"));
		}catch(e){}		
		mo_time=0;
	}
}

function tlogperf(ptime)
{ 
	var ltype="P";
	if(ptime<0)	
	{
		ptime=0;
		ltype="p";
	}
	if(!hastolog("P"))
		return;
	if(mo_tlog==null)
		return;
	if(!mo_tlog)
		mo_tlog=Server.CreateObject(nasterloglognaster);
		
	var lmoment=(new Date()).toLocaleString() ;	
	var lretval=Session("MO_ACTION");
	if(lretval==null || lretval=="")
		lretval=Session("MO_TRANSACTION");
	try{
		mo_tlog.putLogData(lmoment,mo_user,ptime,ltype,Session("MO_SENT_TRANS"),lretval,"",Session("MO_SERVER"));
	}catch(e){}
	Session("MO_SENT_TRANS")=Session("MO_TRANSACTION");
}
function tlogtimer(pcode,ptime)
{ 
	if(mo_tlog==null)
		return;
	if(!mo_tlog)
		mo_tlog=Server.CreateObject(nasterloglognaster);
		
	var lmoment=(new Date()).toLocaleString() ;	
	try{
		mo_tlog.putLogData(lmoment,mo_user,ptime,"M",pcode,"","",Session("MO_SERVER"));
	}catch(e){}
}
function hastolog(peventtype)
{
	if(peventtype==null)
		return true;
	switch(tlevel)
	{
		case "0":  // rien
		case "1":  // erreurs
			return false;
		case "2":  // debut/fin session
			if(peventtype!="D" && peventtype!="F")
				return false;
			break;
		case "3":  // debut/fin session perf
			if(peventtype!="D" && peventtype!="F" && peventtype!="P")
				return false;
			break;
		case "4": //detail
	}
	return true;
}

//---------------------------------------------------------------------- error
function error(mess)
//---------------------------------------------------------------------- 
{
Response.Write ("erreur interceptée par le moniteur<br>"+mess+"<br>");
  Session("MO_ERROR")=mess;
  Server.Transfer(pathfile2("asp/MO_ERROR.ASP"));
  
  
  
  
  var aff="",fic="",b="<br>",r="\r\n",x,fso,f,s;
  fic+="-------------------------------------------------------"+r;
  fic+=Date() +" Utilisateur: "+mo_user+r+r;
  aff+="ERREUR"+b;
  x="Fichier : "+err.File+ "Ligne: "+err.Line;
  aff+=x+b;
  fic+=x+r;
  x="Description: "+err.Description;
  aff+=x+b;
  fic+=x+r;
  x="ASPDescription: "+err.ASPDescription;  
  aff+=x+b;
  fic+=x+r;
  x="Instruction: "+err.Source;
  aff+=x+b;
  fic+=x+r;
  Response.Clear();
  Response.Write(aff);
  
  
  fso = Server.CreateObject("Scripting.FileSystemObject");
  x=Server.MapPath("..\\tmp\\errors.txt");

  Application.Lock ();  

  if (fso.FileExists(x))
  {
    f = fso.GetFile(x);
    ts = f.OpenAsTextStream(8,-2); 
  }
  else
  {   
    ts=fso.CreateTextFile(x);  
  }
  ts.Write(fic);
  ts.Close();
  Application.UnLock ();
  ts=null;
  f=null;
  fso=null;
  %>
  <script language=vbscript>
' force display of page
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx




</script>
<%  
  Response.End();
}


function TraceSession()
{
	Response.Write("\n\r<!--");
	Response.Write("*********Session****************************\n\r");
   fc = new Enumerator(Session.Contents);
   for (; !fc.atEnd(); fc.moveNext())
   {
	  var it = fc.item();
	  var tpe=typeof(Session(it));
	  if ((tpe+'') != 'object' || Session(it)==null)
		Response.Write(it +" = " + Session(it) + "\n\r");
   }
   Response.Write("\n\r************************************************************");
   Response.Write("\n\r-->\n\r");
}
function ClearSession()
{
   fc = new Enumerator(Session.Contents);
   for (; !fc.atEnd(); fc.moveNext())
   {
	  var it = fc.item();
	  Session(it)=null;
   }
}


function TraceStaticObjects()
{
	Response.Write("<!--");
	Response.Write("************************************************************");
   fc = new Enumerator(Application.StaticObjects);
   for (; !fc.atEnd(); fc.moveNext())
   {
	  var it = fc.item();
      Response.Write(it + "= " + Application.StaticObjects(it).Count + "\n\r");
   }
   Response.Write("************************************************************");
   Response.Write("-->");
}
function XmlServiceFactory2(cle)
{
	if(!Session(cle))
		return null;
	var xmls;	
	xmls = Server.CreateObject("XmlService.Document2");	
	xmls.XmlDocument = Session(cle);	
	return xmls;
}	

function MyExecute(page)
{
	//TraceSession();
	Session("MO_TMP")=XSession;
	try{
	Server.Execute(page);
	}catch(e){error("myexecute("+page+")");}
	Session("MO_TMP")=null;
}

function ProfilePerf()
{return;
	var fin = new Date();
	Response.Write("<!-- *** Perf ***  " + (fin - startDate) + " ms écoulées -->");
}
function getAspSessionID()
{
	var h=Request.ServerVariables("HTTP_COOKIE").Item;
	if(!h)
		return null;
	var i=h.indexOf("ASPSESSIONID");
	if(i<0)
		return "";
	var k=h.indexOf(";",i);
	if(k<0)		
		return h.substr(i);
	else
		return h.substring(i,k);
}


function newArray (){ return mo_caprockvect?Server.CreateObject("caprock.Vector"):(new Array()); }
function newArray2(){ return Server.CreateObject("caprock.dictionary"); }
function arrlisttovec(arr)
{
  if(mo_caprockvect)
  {
	var vec=newArray ();
	for(var key in arr)
		vec.Add(key);
  }
  else
  {
	var vec=new Array ();
	for(var key in arr)
		vec.push(key);
  }
  return vec;
}
function setcount(vec,n)    
{
  if(mo_caprockvect)
  {
	var c=vec.length;
	while(c-->n)
	{
		vec(c)=null;
		vec.removeAt(c);
	}
  }
  else
	vec.length=n;
}

function texttrans(t)
{
	Application.Lock ();
	var p=TRANS.content(t);  
    Application.UnLock ();
    return p;
}
function pathfile(f)
{ 
	Application.Lock ();
	var p=MO_FILES(f);  
    Application.UnLock ();
    return p;
}
function pathfile2(f)
{ 
	if(f.substr(f.length-4)==".XML")
		return "/"+site+"/xml/"+f;
	switch(f.substr(0,4))
	{
		case "MDL_": return "/"+site+"/asp/m/"+f; break;
		case "PGL_": return "/"+site+"/asp/v/pgl/"+f; break;
		case "TPL_": return "/"+site+"/asp/v/tpl/"+f; break;
	}
	if(f.indexOf(".")<0)
	  f+=".ASP";
	return "/"+site+"/"+f;
		
}
function ResponseEnd()
{
	if(xxxtran!=0)
	{
		mailtogc("xxxtran= "+xxxtran+descstack());
	}
	Response.End();
}
function mailtogc(mess)
{
  	  var mail = Server.CreateObject("CDONTS.NewMail");
      mail.From="naster@mcsfr.com";
	  mail.To= "bugs@mcsfr.com";   
	  mail.Subject="[gcNASTER]ESPION NASTER utilisateur="+Session("EX_UTILISATEUR")+" serveur="+Session("MO_SERVER");
	  mail.Body= mess;   
      mail.Send();
      mail = null;
}
function logmess(src,mess)
{
  if(mess.substring(0,1)==".") // don't log message
	return;
  var fso = Server.CreateObject("Scripting.FileSystemObject");
  var x="c:\\winnt\\system32\\logfiles\\naster\\errors.txt";
  Application.Lock ();  

  if (fso.FileExists(x))
  {
    var f = fso.GetFile(x);
    var ts = f.OpenAsTextStream(8,-2); // forappending tristatedefault
  }
  else
  {   
    ts=fso.CreateTextFile(x);
  }
  ts.Write("*****logmess= "+Session("ex_utilisateur") + " " + (new Date()) +" ("+src+") "+mess+"\r\n");
  ts.Close();
  Application.UnLock ();
}


function descstack()
{
  var fic,r="\r\n",s,si,ctx;
  
  fic=r+r+"-------------------------------------------------------"+r;
  fic+=Date() +" Utilisateur: "+Session("EX_UTILISATEUR")+r+r;
  fic+="PILE D'APPELS:"+r+r;
  s=Session("MO_CALLSTACK");
  if(s)
  {
	for(var i=0;i<s.length;i++)
	{ 
	  try{si=s(i);}catch(e){si=s[i];} //caprock
	  getpos(si);
	  fic+="transaction: "+si("name")+" ligne: "+stline+" "+stinstr+r;
	}
  }
  ctx=Session("MO_CTX");
  getpos(ctx);
  fic+="transaction: "+stname+" ligne: "+stline+" "+stinstr+r;
  return fic;
}

function getpos(ctx)
{
 var prg,ptr,n,p1;
 
 stname="";
 stinstr="";
 stline=0;
 
 if(ctx==null)
   return;

 stname=ctx("name");
 stline=1;
 prg=ctx("prg");
 ptr=ctx("ptr")
 n=0;
 p1=0;
 while(true)
 {
   n=prg.indexOf("\n",n);
   if(n<0 || n>=ptr)
     break;
   stline++;
   n++;
   p1=n;
 }

 n=p1;
 while(true)
 {
   n=prg.indexOf(";",n);
   if(n<0)
   {
     stinstr=prg.substr(p1);
     break;
   }
   if(n>=ptr)
   {
     stinstr=prg.substring(p1,n+1);
     break;
   }
   p1=n;
   n++;
 }
}

%>

