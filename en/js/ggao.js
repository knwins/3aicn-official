// ------ ����ȫ�ֱ��� 
       var theNewsNum; 
    var theAddNum; 
    var totalNum; 
    var CurrentPosion=0; 
       var theCurrentNews; 
       var theCurrentLength; 
       var theNewsText; 
       var theTargetLink; 
       var theCharacterTimeout; 
       var theNewsTimeout; 
       var theBrowserVersion; 
       var theWidgetOne; 
       var theWidgetTwo; 
       var theSpaceFiller; 
       var theLeadString; 
       var theNewsState; 
       function startTicker() 
       {                
// ------ ���ó�ʼ��ֵ 
          theCharacterTimeout = 50;//�ַ����ʱ�� 
          theNewsTimeout     = 3000;//���ż��ʱ�� 
          theWidgetOne        =  "_";//����ǰ���±��1 
          theWidgetTwo        =  "-";//����ǰ���±�� 
          theNewsState       = 1; 
          //theNewsNum        = document.body.children.incoming.children.NewsNum.innerText;//���������� 
          //add by lin 
     theNewsNum        = document.getElementById("incoming").children.AllNews.children.length;//���������� 
     theAddNum        = document.getElementById("incoming").children.AddNews.children.length;//�������� 
     totalNum   =theNewsNum+theAddNum; 
     theCurrentNews     = 0; 
          theCurrentLength    = 0; 
          theLeadString       = " "; 
          theSpaceFiller      = " "; 
          runTheTicker(); 
       } 
// --- �������� 
       function runTheTicker() 
       { 
          if(theNewsState == 1) 
          { 
            if(CurrentPosion<theNewsNum){  
          setupNextNews(); 
            } 
      else{ 
          setupAddNews(); 
      } 
      CurrentPosion++; 
      if(CurrentPosion>=totalNum||CurrentPosion>=5) CurrentPosion=0;  //�������������5�� 
     } 
          if(theCurrentLength != theNewsText.length) 
          { 
             drawNews(); 
          } 
          else 
          { 
             closeOutNews(); 
          } 
       } 
// --- ��ת��һ������ 
       function setupNextNews() 
       { 
          theNewsState = 0; 
     theCurrentNews = theCurrentNews % theNewsNum;      
          theNewsText = document.getElementById("AllNews").children[theCurrentNews].children.Summary.innerText; 
          theTargetLink = document.getElementById("AllNews").children[theCurrentNews].children.NewsLink.innerText;           
          theCurrentLength = 0; 
          document.all.hottext.href = theTargetLink; 
          theCurrentNews++; 
    } 
       function setupAddNews() 
       { 
          theNewsState = 0; 
     theCurrentNews = theCurrentNews % theAddNum;      
          theNewsText = document.getElementById("incoming").children.AddNews.children

[theCurrentNews].children.Summary.innerText; 
          theTargetLink = document.getElementById("incoming").children.AddNews.children

[theCurrentNews].children.NewsLink.innerText;           
          theCurrentLength = 0; 
          document.all.hottext.href = theTargetLink; 
          theCurrentNews++; 
    }     
// --- �������� 
       function drawNews() 
       { 
          var myWidget;        
          if((theCurrentLength % 2) == 1) 
          { 
             myWidget = theWidgetOne; 
          } 
          else 
          { 
             myWidget = theWidgetTwo; 
          } 
          document.all.hottext.innerHTML = theLeadString + theNewsText.substring(0,theCurrentLength) + myWidget + 

theSpaceFiller; 
          theCurrentLength++; 
          setTimeout("runTheTicker()", theCharacterTimeout); 
       } 
// --- ��������ѭ�� 
       function closeOutNews() 
       { 
          document.all.hottext.innerHTML = theLeadString + theNewsText + theSpaceFiller; 
          theNewsState = 1; 
          setTimeout("runTheTicker()", theNewsTimeout); 
       }       
window.onload=startTicker;         
//--> 