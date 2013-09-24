
#include <iostream>
#include <windows.h>
#include <mshtml.h>
#include <oleacc.h>
#include <atlbase.h>
#include <atlbase.h> //��Ҫ��װATL�� 
/****************************************************************************
Ѱ��ָ���������Ӵ��ھ��
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
��һ�����ھ����ȡIHTMLDocument2�ӿ�
ʹ�����Ҫ����Release
����Ҳ����ӿڣ�����NULL
ԭ��
������ϵͳ��װ��Microsoft ��������ܣ�MSAA�����������������������
������"Internet Explorer_Server"������WM_HTML_GETOBJECT��Ϣ������Ϣ���صĽ��
��Ϊһ���������ݸ�MSAA����ObjectFromLresult���Ӷ���ȡIHTMLDocument2 �ӿڡ�

****************************************************************************/  

void GetIHTMLDocument2Interface(HWND BrowserWnd)
{
 CoInitialize(NULL);
 HRESULT hr;
 // Explicitly load MSAA so we know if it's installed
 HINSTANCE hInst = ::LoadLibrary( _T("OLEACC.DLL") );
 if ( hInst )
 {
  LRESULT lRes; //SendMessageTimeout��ķ���ֵ�����ں���pfObjectFromLresult�ĵ�1������
  UINT nMsg = ::RegisterWindowMessage( _T("WM_HTML_GETOBJECT") );
  ::SendMessageTimeout( BrowserWnd, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes );
  //��ȡ����pfObjectFromLresult
  LPFNOBJECTFROMLRESULT pfObjectFromLresult = (LPFNOBJECTFROMLRESULT)::GetProcAddress( hInst, "ObjectFromLresult");
  if ( pfObjectFromLresult  )
  {
   CComPtr<IHTMLDocument2> spDoc;
   hr = (*pfObjectFromLresult)( lRes, IID_IHTMLDocument, 0, (void**)&spDoc );
   if ( SUCCEEDED(hr) )
   {
    //��ȡ�ĵ��ӿ�
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
