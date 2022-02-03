#include <windows.h>
#include <mmsystem.h>
#include <stdlib> 
#include "resource.h"
#include <commctrl.h>
#include <iostream>
#include <sstream>
#include <stdio>
#include <math.h>
#include <string>


using namespace std;


LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK FrameProc(HWND hWnd, UINT, WPARAM , LPARAM);
LPCSTR lpszAppName = "KineVideo";
LPCSTR lpszTitle   = "KineVideo";
char s_title[200];

HWND hmain;
HWND cmdBack;
HWND cmdNull;
HWND cmdDrawF;
HWND t_function;
HWND t_start;
HWND t_end;
HWND l_x;
HWND p_draw;
HWND cmdCali;
HWND cmdNullT;
HWND cmdFor;
HWND textCali;
HWND l_Cali;
HWND l_time;
HWND hvpanel;
HWND hinfo;
HWND c_slide;
HCURSOR hcross;
HFONT hFont;
HFONT hFont2;
HFONT hFont3;
HICON i_null;
HICON i_nullt;
HICON i_cross1;
HICON i_cross2;

#define BACK      1
#define FOR		  2
#define JUMP	  3
#define STATI	  4
#define ZERO	  5
#define CALI	  6
#define CALITXT	  7
#define ZEROT	  8
#define DRAW	  9
#pragma resource "kinevideo.res"


LPSTR mciCmd(LPSTR command);
void setZoom(bool b_refresh);
void stepAVI(int i_steps,bool b_jump);
LPSTR SaveFileBox(HWND hwnd,bool b_excel);
void OpenFileBox(HWND hwnd);
void getAviInfo();
void updateAviInfo();
bool LoadAvi(LPCSTR s_videofile);
bool ExportData();
void DrawFunction();
float calcScale();
void GetPara();

HWND CreateToolTip(HWND toolID, HWND hDlg, PTSTR pszText,HINSTANCE hInstance);
struct s_info_avi
{
 int i_fps;
 int i_width;
 int i_height;
 int i_frames;
 int cur_pos;
};
struct s_info_avi info_avi;
int* x_values;
int* y_values;
int x_s[4];
int y_s[4];
int i_markmode;
float f_calilength;
bool b_already;
int i_screen_w;
int i_screen_h;
bool b_null;
bool b_cali;
LPSTR s_XMLHEADER;
char s_filename[256];
char szFileName[256];
char s_error_msg[300];
int i_tnull;
float f_alpha;
bool b_zoom;
bool b_dfunc;
float f_para[3];

int APIENTRY WinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance, PSTR szCmdLine, int iCmdShow)
{

   

   HWND       hWnd;
   MSG        msg;
   WNDCLASSEX   wc;

   wc.cbSize        =  sizeof(WNDCLASSEX);
   wc.style         =  CS_HREDRAW | CS_VREDRAW | CS_NOCLOSE; 
   wc.lpfnWndProc   =  WndProc;
   wc.cbClsExtra    =  0;
   wc.cbWndExtra    =  0;
   wc.hInstance     =  hInstance;
   wc.hCursor       =  LoadCursor(NULL,IDC_ARROW);
   wc.hIcon         =  (HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(ID_APP), IMAGE_ICON, 32, 32, 0);
   wc.hbrBackground =  (HBRUSH)(COLOR_BTNFACE + 1);
   wc.lpszClassName =  lpszAppName;
   wc.lpszMenuName  =  MAKEINTRESOURCE(IDR_MENU1);
   wc.hIconSm       =  (HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(ID_APP), IMAGE_ICON, 16, 16, 0);

   if( RegisterClassEx(&wc) == 0)
      return 0;

	HDC l_dc;
	l_dc = GetDC(NULL);
    i_screen_w = GetDeviceCaps(l_dc, HORZRES);
    i_screen_h = GetDeviceCaps(l_dc, VERTRES);
	  
   hWnd = CreateWindowEx(WS_EX_CLIENTEDGE,
                         lpszAppName,
                         lpszTitle,
                         WS_OVERLAPPEDWINDOW,
                         0,
                         0,
                         i_screen_w,
                         i_screen_h,
                         NULL,
                         NULL,
                         hInstance,
                         NULL);

   if( hWnd == NULL)
      return 0;

	  
	  
   ShowWindow(hWnd, iCmdShow);
   UpdateWindow(hWnd);
 
   hmain = hWnd;
   hFont2=CreateFont (12, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, \
		  OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, \
		  DEFAULT_PITCH | FF_SWISS, "Arial"); 
   hFont3=CreateFont (11, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, \
		  OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, \
		  DEFAULT_PITCH | FF_SWISS, "Arial"); 
   hFont=CreateFont (14, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, \
		  OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, \
		  DEFAULT_PITCH | FF_SWISS, "Arial");
   
   cmdFor =   CreateWindow("button","Vor",WS_CHILD|WS_VISIBLE|
                          BS_DEFPUSHBUTTON|BS_ICON, 30, 0, 30, 30,
                          hWnd, (HMENU)FOR, hInstance, NULL);
   cmdBack =   CreateWindow("button","",WS_CHILD|WS_VISIBLE|
                          BS_DEFPUSHBUTTON|BS_ICON, 0, 0, 30, 30,
                          hWnd, (HMENU)BACK, hInstance, NULL);
   cmdNullT =   CreateWindow("button","Nullpunkt",WS_CHILD|WS_VISIBLE|
                          BS_DEFPUSHBUTTON|BS_ICON, 170, 0, 30, 30,
                          hWnd, (HMENU)ZEROT, hInstance, NULL);
   cmdNull =   CreateWindow("button","Nullpunkt",WS_CHILD|WS_VISIBLE|
                          BS_DEFPUSHBUTTON|BS_ICON, 200, 0, 30, 30,
                          hWnd, (HMENU)ZERO, hInstance, NULL);
	cmdCali =   CreateWindow("button","Kalibrieren",WS_CHILD|WS_VISIBLE|
                          BS_DEFPUSHBUTTON|BS_ICON, 230, 0, 30, 30,
                          hWnd, (HMENU)CALI, hInstance, NULL);
    hinfo = CreateWindow("static","",WS_CHILD|WS_VISIBLE, 480, 0, 320, 30, hWnd, NULL, hInstance, NULL);
	
	l_Cali = CreateWindow("static","Kalibrierungsstrecke (in m):\nZeitnullpunkt: Bild 1",WS_CHILD|WS_VISIBLE, 270, 0, 135, 30,hWnd, NULL, hInstance, NULL);
	l_time = CreateWindow("static","",WS_CHILD|WS_VISIBLE, 390, 16, 70, 12,hWnd, NULL, hInstance, NULL);

	
	t_function = CreateWindow("edit",NULL,WS_BORDER|WS_CHILD|WS_VISIBLE, 480, 0, 200, 15, hWnd, NULL, hInstance, NULL);
	cmdDrawF = CreateWindow("button","Zeichnen",WS_CHILD|WS_VISIBLE|BS_DEFPUSHBUTTON, 700, 0, 60, 30,hWnd, (HMENU)DRAW, hInstance, NULL);
    t_start = CreateWindow("edit",NULL,WS_BORDER|WS_CHILD|WS_VISIBLE, 560, 15, 30, 15, hWnd, NULL, hInstance, NULL);
	t_end = CreateWindow("edit",NULL,WS_BORDER|WS_CHILD|WS_VISIBLE, 600, 15, 30, 15, hWnd, NULL, hInstance, NULL);
	l_x = CreateWindow("static","Intervall x-Werte:",WS_CHILD|WS_VISIBLE, 480, 15, 80, 15,hWnd, NULL, hInstance, NULL);
	ShowWindow(t_function,SW_HIDE);
	ShowWindow(t_start,SW_HIDE);
	ShowWindow(t_end,SW_HIDE);
	ShowWindow(l_x,SW_HIDE);
	ShowWindow(cmdDrawF,SW_HIDE);
	b_dfunc = false;
	
	hvpanel =   CreateWindow("static","",WS_CHILD|WS_VISIBLE, 0, 30, 960, 536, hWnd, (HMENU)STATI, hInstance, NULL);
	SetWindowLong(hvpanel, GWL_WNDPROC, (LONG) FrameProc); 
	c_slide = CreateWindow (TRACKBAR_CLASS, "",WS_CHILD | WS_VISIBLE | TBS_AUTOTICKS,60, 0, 110, 30,hWnd, (HMENU)JUMP, hInstance, NULL);
	textCali = CreateWindow("edit",NULL,WS_BORDER|WS_CHILD|WS_VISIBLE, 410, 0, 50, 16, hWnd, NULL, hInstance, NULL);
	//Neue Fonts setzen

	SendMessage (l_time, WM_SETFONT, WPARAM (hFont), TRUE);
	SendMessage (hinfo, WM_SETFONT, WPARAM (hFont), TRUE);
	SendMessage (textCali, WM_SETFONT, WPARAM (hFont2), TRUE);
	SendMessage (t_function, WM_SETFONT, WPARAM (hFont3), TRUE);
	SendMessage (t_start, WM_SETFONT, WPARAM (hFont3), TRUE);
	SendMessage (t_end, WM_SETFONT, WPARAM (hFont3), TRUE);
	SendMessage (l_x, WM_SETFONT, WPARAM (hFont2), TRUE);
	SendMessage (cmdDrawF, WM_SETFONT, WPARAM (hFont2), TRUE);
	SendMessage (l_Cali, WM_SETFONT, WPARAM (hFont), TRUE);
	//Schaltflächen-Bilder setzen
	i_null = (HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(IDI_NULL), IMAGE_ICON, 25,25, 0);
	i_nullt = (HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(IDI_NULLT), IMAGE_ICON, 25,25, 0);
	i_cross1 = (HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(ID_CROSS1), IMAGE_ICON, 32,32, 0);
	i_cross2 = (HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(ID_CROSS2), IMAGE_ICON, 32,32, 0);
	SendMessage (cmdBack, BM_SETIMAGE,IMAGE_ICON,(LPARAM)(HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(IDI_BACK), IMAGE_ICON, 25,25, 0));
	SendMessage (cmdFor, BM_SETIMAGE,IMAGE_ICON,(LPARAM)(HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(IDI_FOR), IMAGE_ICON, 25,25, 0));
	SendMessage (cmdNull, BM_SETIMAGE,IMAGE_ICON,(LPARAM)i_null);
	SendMessage (cmdNullT, BM_SETIMAGE,IMAGE_ICON,(LPARAM)i_nullt);
	SendMessage (cmdCali, BM_SETIMAGE,IMAGE_ICON,(LPARAM)(HICON)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(IDI_CALI), IMAGE_ICON, 25,25, 0));
	//ToolTipText
	CreateToolTip(cmdFor,hWnd,"Ein Bild vor",hInstance);
	CreateToolTip(cmdBack,hWnd,"Ein Bild zurück",hInstance);
	CreateToolTip(cmdNull,hWnd,"Koordinaten-Ursprung festlegen",hInstance);
	CreateToolTip(cmdNullT,hWnd,"Zeitnullpunkt festlegen",hInstance);
	CreateToolTip(cmdCali,hWnd,"Kalibrierungsstrecke festlegen",hInstance);
	
	//Koordinatensystem initalisieren
	x_s[0]=0;
	y_s[0]=0;
    i_markmode=0; 
	//Länge der Kalibrierungsstrecke in Meter
	f_calilength = 1.00;
	b_already = false;
	info_avi.i_fps = 25;
	b_zoom = false;
	HMENU hMenu = GetMenu(hWnd);
	if (b_zoom)
	{
	 CheckMenuItem(hMenu, (UINT) ID_ZOOM, MF_BYCOMMAND | MF_CHECKED); 
	}
	else
	{
	 CheckMenuItem(hMenu, (UINT) ID_ZOOM, MF_BYCOMMAND | MF_UNCHECKED); 
	}
    
	SetWindowText(textCali,"1.00");
	SetWindowText(t_start,"0");
	SetWindowText(t_end,"306");

    
   hcross = (HCURSOR)LoadImage(GetModuleHandle(NULL),MAKEINTRESOURCE(ID_CURSOR), IMAGE_ICON, 32, 32, 0);
   SetClassLong(hvpanel,GCL_HCURSOR,(LONG) hcross); 
   
   //XML-HEADER setzen
    s_XMLHEADER="<?xml version=\"1.0\"?>\n<?mso-application progid=\"Excel.Sheet\"?>\n";
    s_XMLHEADER=strcat(s_XMLHEADER,"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
	s_XMLHEADER=strcat(s_XMLHEADER,"xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
	s_XMLHEADER=strcat(s_XMLHEADER,"xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n");
	s_XMLHEADER=strcat(s_XMLHEADER,"xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
	s_XMLHEADER=strcat(s_XMLHEADER, "xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");

   while (GetMessage(&msg, NULL, 0, 0) > 0)
   {
      TranslateMessage(&msg);
      DispatchMessage(&msg);
   }
   return msg.wParam;
}
void GetPara()
{
  char s_answer[100];
 char s_answer2[100];
 GetWindowText(t_function,s_answer,sizeof(s_answer));
 

 for (int k=0;k<3;k++)
 {
  string s_ans(s_answer);
  //Kommas in Punkte umwandeln
  s_answer[s_ans.find(",")]='.';
 }
 string s_ans(s_answer);
 // Leerzeichen entfernen
 int counts;
 counts = 0;
 for (int k=0;k<s_ans.length();k++)
 {
  if (s_answer[k]!=' ')
   s_answer2[counts++]=s_answer[k];
 } 
 sscanf(s_answer2,"y=%fx2%fx%f",&f_para[2],&f_para[1],&f_para[0]);
}
void DrawFunction()
{
 HDC hdc; 
 HPEN hPen; 
 int i_samples;
 float f_int[2];
 float f_ox;
 float f_oy;
 float f_x;
 float f_y; 
 float i_tx;
 float i_ty;

 float f_scale;



 hdc = GetDC(hvpanel); 
 hPen = CreatePen(PS_INSIDEFRAME, 3, RGB(0,255,0));
 SelectObject(hdc, hPen);
 
 f_scale = calcScale();
 
 
 
 i_samples = 1000;
 
 char sBuffer[10];
 GetWindowText(t_start,sBuffer,sizeof(sBuffer));
 string s_ans(sBuffer);
  //Kommas in Punkte umwandeln
 sBuffer[s_ans.find(",")]='.';
 f_int[0] =atof(sBuffer); 
 
 GetWindowText(t_end,sBuffer,sizeof(sBuffer));
 s_ans = sBuffer;
  //Kommas in Punkte umwandeln
 sBuffer[s_ans.find(",")]='.';
 
 f_int[1] =atof(sBuffer);

 f_ox = f_int[0];
 f_oy = f_para[2]*f_ox*f_ox+f_para[1]*f_ox+f_para[0];
 
  f_x = f_ox*cos(f_alpha)+f_oy*sin(f_alpha);
  f_y = -f_ox*sin(f_alpha)+f_oy*cos(f_alpha); 
 
  i_tx = (int)(f_x/f_scale)+x_s[0];
  i_ty = -(int)(f_y/f_scale)+y_s[0];

  MoveToEx(hdc, i_tx, i_ty, NULL);


 for (float j=0;j<float(i_samples);j++)
 {
   /*x_p[0] = (x_values[j]-x_s[0])*f_scale;
   y_p[0] = -(y_values[j]-y_s[0])*f_scale;
   //Drehen
   x_p[1] = x_p[0]*cos(f_alpha)-y_p[0]*sin(f_alpha);
   y_p[1] = x_p[0]*sin(f_alpha)+y_p[0]*cos(f_alpha);*/
  f_ox = (f_int[1]-f_int[0])*(j/(float)i_samples)+f_int[0];
  f_oy = f_para[2]*f_ox*f_ox+f_para[1]*f_ox+f_para[0];


  f_x = f_ox*cos(f_alpha)+f_oy*sin(f_alpha);
  f_y = -f_ox*sin(f_alpha)+f_oy*cos(f_alpha); 


  i_tx = (int)(f_x/f_scale)+x_s[0];
  i_ty = -(int)(f_y/f_scale)+y_s[0];
  
  LineTo (hdc, i_tx, i_ty); 

 }
 ReleaseDC(hvpanel,hdc);

 

}
void UpdateAviInfo()
{
 char cmd[300];
 char s_cali[300];
 char s_null[300];
 if (b_null==false)
 {
  sprintf(s_null,"Nullpunkt noch nicht festgelegt!");
 }
 else
 {
  sprintf(s_null,"Nullpunkt festgelegt!");
 }
 if (b_cali==false)
 {
  sprintf(s_cali,"Kalibrierung noch nicht durchgeführt!");
 }
 else
 {
  sprintf(s_cali,"Kalibrierung durchgeführt!");
 }
 sprintf(cmd, "Bildrate: %i fps \t\t%s\nBildposition: %i/%i\t%s",info_avi.i_fps,s_null,info_avi.cur_pos+1,info_avi.i_frames,s_cali);
 
  SetWindowText(hinfo,cmd);
  sprintf(cmd, "Kalibrierungsstrecke (in m):\nZeitnullpunkt: Bild %i",i_tnull+1);
  SetWindowText(l_Cali,cmd);
   
  sprintf(cmd, "Zeit: %.2fs",(float)(info_avi.cur_pos-i_tnull)/(float)info_avi.i_fps);
   
  SetWindowText(l_time,cmd);
}
void getAviInfo()
{
   LPSTR s_result;
   char cmd[300];
   int i_pos;
   float f_frame;
   
   wsprintf(cmd, "status myAVI nominal frame rate");
   s_result = mciCmd(cmd);
    
   info_avi.i_fps = ceil(atof(s_result)/1000.0);
   wsprintf(cmd, "status myAVI length");
   s_result = mciCmd(cmd);
	
   info_avi.i_frames = atoi(s_result);
   
   wsprintf(cmd, "where myAVI destination");
   s_result = mciCmd(cmd);

   string str_result(s_result);
   i_pos = str_result.find(" ",0);
   i_pos = str_result.find(" ",i_pos+1);
   str_result = str_result.substr(i_pos+1);
   i_pos = str_result.find(" ",0);
   info_avi.i_height = atoi(str_result.substr(i_pos+1).c_str());
   info_avi.i_width = atoi(str_result.substr(0,i_pos).c_str());
   info_avi.cur_pos = 0;

   
   //Trackbar setzen
   SendMessage(c_slide, TBM_SETRANGEMIN, true, 0);
   SendMessage(c_slide, TBM_SETRANGEMAX, true, info_avi.i_frames-1);

   //Array initialisieren:
   x_values = new int[info_avi.i_frames];
   y_values = new int[info_avi.i_frames];
   for(int j=0;j<info_avi.i_frames;j++)
   {
    x_values[j] = 0;
	y_values[j] = 0;
   }
 
   UpdateAviInfo();


}
bool LoadAvi(LPCSTR s_videofile)
{
   b_null = false;
   b_cali = false;
   
   //Zeitnullpunkt setzen:
   i_tnull = 0;
   //Winkel setzen:
   f_alpha = 0;
   
   char sFile[300];
   GetShortPathName(s_videofile,sFile,sizeof(sFile));
   char cmd[300];
   if (b_already == true)
   {
    wsprintf(cmd, "close myAVI");
    mciCmd(cmd);
   }
   
   wsprintf(cmd, "open \"%s\" type avivideo alias myAVI", sFile);
   string s_error(mciCmd(cmd));

   //MessageBox(NULL,s_error.c_str(),"Fehler",0);
   if (s_error == "failed")
   {
    MessageBox(NULL,s_error_msg,"Fehler",0);
    return false;
   }

   wsprintf(cmd, "window myAVI handle %lu" ,hvpanel);
   mciCmd(cmd);

   getAviInfo();
    
   setZoom(false);

   wsprintf(cmd, "set myAVI seek exactly on");
   mciCmd(cmd);  

   wsprintf(cmd, "set myAVI time format frames");
   mciCmd(cmd);   

   wsprintf(cmd, "play myAVI from 0");
   string s_error2(mciCmd(cmd));

   //MessageBox(NULL,s_error2.c_str(),"Fehler",0);
   if (s_error2 == "failed\0")
   {
    MessageBox(NULL,s_error_msg,"Fehler",MB_ICONERROR);
    return false;
   }

   wsprintf(cmd, "pause myAVI");
   mciCmd(cmd);  
   wsprintf(cmd, "seek myAVI to 0");
   mciCmd(cmd);  
   SendMessage(c_slide, TBM_SETPOS, true, 0);
   
   b_already = true;
   return true;
}
void setZoom(bool b_refresh)
{
   char cmd[300];
   //Video auf Fenster anpassen
   float f_ratio;
   float f_screen;

   if (b_zoom == true)
   {
    f_ratio = (float)info_avi.i_height/(float)info_avi.i_width;
    f_screen = (float)(i_screen_h-80)/(float)i_screen_w;
   

    if (f_ratio < f_screen)
    {
     wsprintf(cmd, "put myAVI destination at 0 0 %i %i",i_screen_w,(int)(i_screen_w*f_ratio));
	 //MessageBox(NULL,cmd,"1",0);
	 SetWindowPos(hvpanel,NULL,0,30,i_screen_w,(int)(i_screen_w*f_ratio),SWP_SHOWWINDOW);
    }
    else
    {
     wsprintf(cmd, "put myAVI destination at 0 0 %i %i",(int)((i_screen_h-80)/f_ratio),i_screen_h-80);
	 //MessageBox(NULL,cmd,"2",0);
	 SetWindowPos(hvpanel,NULL,0,30,(int)((i_screen_h-80)/f_ratio),i_screen_h-80,SWP_SHOWWINDOW);
    }
	mciCmd(cmd);
   }
   else
   {
    wsprintf(cmd, "put myAVI destination at 0 0 %i %i",info_avi.i_width,info_avi.i_height);
    mciCmd(cmd);
   }
   if (b_refresh)
   {
    stepAVI(1,false);
	stepAVI(-1,false);
   }
}
LPSTR mciCmd(LPSTR command)
{
   long l_error;
   char s_Return[256];
   l_error = mciSendString(command, s_Return, 256, 0);
   //MessageBox(NULL,command,"",0);
   if (l_error != 0)
   {
    char s_error[300];
    mciGetErrorString(l_error,s_error,sizeof(s_error));
	//MessageBox(NULL, command, "cmd", 0);
	
	//MessageBox(NULL, s_error, "Fehler", 0);
	//sprintf(s_error,"%f",(float)l_error);
	//MessageBox(NULL, s_error, "Fehler", 0);
	
	ltoa(l_error, s_error,sizeof(300));
	
    switch(l_error)
	{
	 case MCIERR_CANNOT_LOAD_DRIVER:
	 {
	  sprintf(s_error_msg,"Es gibt Probleme mit dem Codec! Vermutlich ist nicht\nder Richtige installiert!");
	  return "failed";
	 }
	 case MCIERR_DEVICE_TYPE_REQUIRED:
	 {
	  sprintf(s_error_msg,"Es gibt Probleme mit dem Codec! Vermutlich ist nicht\nder Richtige installiert!");
	  return "failed";
	 }
	 case MCIERR_DRIVER:
	 {
	  sprintf(s_error_msg,"Es gibt Probleme mit dem Codec! Vermutlich ist nicht\nder Richtige installiert!");
	  return "failed";
	 }
	 case MCIERR_DRIVER_INTERNAL:
	 { 
	  sprintf(s_error_msg,"Es gibt Probleme mit dem Codec! Vermutlich ist nicht\nder Richtige installiert!");
	  return "failed";
	 }
	{
	 
	}
   }
  } 
  return s_Return;
}
LPSTR SaveFileBox(HWND hwnd,bool b_excel)
{
   OPENFILENAME ofn;
   char szSaveName[256]="";
   ZeroMemory(&ofn, sizeof(ofn));

   ofn.lStructSize = sizeof(ofn);
   ofn.hwndOwner = hwnd;

   ofn.lpstrFile = szSaveName;
   ofn.nMaxFile = MAX_PATH;
   ofn.Flags = OFN_EXPLORER;
    
   if(b_excel)
   {
    ofn.lpstrFilter = "XML-Datei (*.xml)\0*.xml\0";
	ofn.lpstrDefExt = "xml";
   }
   else
   {
	ofn.lpstrFilter = "Textdatei (*.txt)\0*.txt\0";   
	ofn.lpstrDefExt = "txt";
   }
   
   if(GetSaveFileName(&ofn))
      {
         return szSaveName;
      }
	  {
	    return "zero";
	  }
}
void OpenFileBox(HWND hwnd)
{
   OPENFILENAME ofn;
   

   ZeroMemory(&ofn, sizeof(ofn));

   ofn.lStructSize = sizeof(ofn);
   ofn.hwndOwner = hwnd;
   ofn.lpstrFilter = "AVI-Datei (*.avi)\0*.avi\0"
                     "Alle Dateien (*.*)\0*.*\0";
   ofn.lpstrFile = szFileName;
   ofn.nMaxFile = MAX_PATH;
   ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST ;
   ofn.lpstrDefExt = "avi";
 
   
   
   if(GetOpenFileName(&ofn))
      {
	     string s_temp(szFileName);
		 sprintf(s_filename,"%s",s_temp.substr(s_temp.rfind("\\")+1).c_str());
         LoadAvi(szFileName);
      }
}
void stepAVI(int i_steps,bool b_jump)
{
  char * cmd;
  cmd = (char *) malloc(300);

 if ((i_steps == 0) && (b_jump == false))
 {
  sprintf(cmd, "update myAVI hdc %u" ,(unsigned int) GetDC(hvpanel));
 }
 else
 {
  if (b_jump == false)
  {
   sprintf(cmd, "step myAVI by %i" ,i_steps);
   info_avi.cur_pos += i_steps;
  }
  else
  {
   sprintf(cmd, "seek myAVI to %i" ,i_steps);
   info_avi.cur_pos = i_steps;
  }
 }
 mciCmd(cmd); 
 //SendMessage(c_slide, TBM_SETPOS, true, info_avi.cur_pos );
 UpdateAviInfo();
 free(cmd);

}
float calcScale()
{
 int i_calipixel;
 char sBuffer[10];
 //Länge der Kalibrierungsstrecke in Pixel
 i_calipixel = sqrt(pow(x_s[2]-x_s[1],2)+pow(y_s[2]-y_s[1],2));
 if (i_calipixel == 0)
	i_calipixel = 1;
 GetWindowText(textCali,sBuffer,sizeof(sBuffer));
 string s_ans(sBuffer);
  //Kommas in Punkte umwandeln
 sBuffer[s_ans.find(",")]='.';
  f_calilength = atof(sBuffer);
 if (f_calilength == 0.0)
	f_calilength = 1.0;
 //Skalierung berechnen
 return f_calilength/float(i_calipixel);
}
bool ExportData(bool b_excel)
{
 FILE *exportfile;

 float f_scale;
 float x_p[2];
 float y_p[2];

 string s_savefile;


 f_scale = calcScale();


 
 //in normale Text-Datei schreiben
 if (!b_excel)
 {
  s_savefile = SaveFileBox(hmain,false);
  if (s_savefile == "zero\0")
  {
	return false;
  }
  exportfile = fopen(s_savefile.c_str(),"w");
  if (exportfile == NULL)
  {
   MessageBox(NULL,"Zieldatei kann nicht geöffnet werden!","Fehler",MB_ICONERROR);
   return false;
  }
  for(int j=0;j<info_avi.i_frames;j++)
  {
   x_p[0] = (x_values[j]-x_s[0])*f_scale;
   y_p[0]= -(y_values[j]-y_s[0])*f_scale;
   //Drehen
   x_p[1] = x_p[0]*cos(f_alpha)-y_p[0]*sin(f_alpha);
   y_p[1] = x_p[0]*sin(f_alpha)+y_p[0]*cos(f_alpha);   
   fprintf(exportfile,"%f\t%f\t%f\n",(float)(j-i_tnull)/(float)info_avi.i_fps,x_p[1],y_p[1]);
  }
  fclose(exportfile);
 }
 //in Excel-Datei schreiben
 else
 {
  char s_temp[500];
  
  s_savefile = SaveFileBox(hmain,true);
  if (s_savefile == "zero")
  {	
  return false;
  }

  exportfile = fopen(s_savefile.c_str(),"w");
  if (exportfile == NULL)
  {
   MessageBox(NULL,"Zieldatei kann nicht geöffnet werden!","Fehler",MB_ICONERROR);
   return false;
  }
  //Header schreiben
  fprintf(exportfile,"%s",s_XMLHEADER);
  //Tableheader schreiben
  sprintf(s_temp,"<Worksheet ss:Name=\"Messdaten\">\n");
  fprintf(exportfile,"%s",s_temp);
  sprintf(s_temp,"<Table ss:ExpandedColumnCount=\"%i\" ss:ExpandedRowCount=\"%i\" x:FullColumns=\"1\"\n x:FullRows=\"1\" ss:DefaultColumnWidth=\"60\" ss:DefaultRowHeight=\"15\">\n",3,info_avi.i_frames+1);
  fprintf(exportfile,"%s",s_temp);  
  sprintf(s_temp,"<Row>\n\t<Cell><Data ss:Type=\"String\">t/s</Data></Cell>\n\t<Cell><Data ss:Type=\"String\">x/m</Data></Cell>\n\t<Cell><Data ss:Type=\"String\">y/m</Data></Cell>\n</Row>\n");
  fprintf(exportfile,"%s",s_temp); 
  for(int j=0;j<info_avi.i_frames;j++)
  {
   if((x_values[j] !=0) && (y_values[j] !=0))
   {
    x_p[0] = (x_values[j]-x_s[0])*f_scale;
    y_p[0] = -(y_values[j]-y_s[0])*f_scale;
    //Drehen
    x_p[1] = x_p[0]*cos(f_alpha)-y_p[0]*sin(f_alpha);
    y_p[1] = x_p[0]*sin(f_alpha)+y_p[0]*cos(f_alpha);
    sprintf(s_temp,"<Row>\n\t<Cell><Data ss:Type=\"Number\">%f</Data></Cell>\n\t<Cell><Data ss:Type=\"Number\">%f</Data></Cell>\n\t<Cell><Data ss:Type=\"Number\">%f</Data></Cell>\n</Row>\n",(float)(j-i_tnull)/(float)info_avi.i_fps,x_p[1],y_p[1]);
    fprintf(exportfile,"%s",s_temp); 
    if ((j+1) == info_avi.i_frames)
    {
     char s_s[30];
     sprintf(s_s,"%i",(int)x_p[1]);
     SetWindowText(t_end,s_s);
    }
   }
  }
  sprintf(s_temp,"</Table>\n</Worksheet>\n</Workbook>");
  fprintf(exportfile,"%s",s_temp); 
  fclose(exportfile);

 }

 return true;
}
LRESULT CALLBACK FrameProc(HWND hWnd, UINT umsg, WPARAM wParam, LPARAM lParam)
{
if (b_already)
{
  switch (umsg)
   { 

	break;
    case WM_MOUSEMOVE:
	{
	 int x, y;
	 
	 x = LOWORD( lParam );
     y = HIWORD( lParam );		
	 sprintf(s_title,"KineVideo - %s - (%i|%i) ",s_filename,x,y);
	 SetWindowText(hmain,s_title);
	 return 0;
	}
    case WM_LBUTTONUP:
	{
		 int x, y;
		 char msg[300];
         x = LOWORD( lParam );
         y = HIWORD( lParam );	
	    if (i_markmode == 0)
		 {
          x_values[info_avi.cur_pos] = x;
		  y_values[info_avi.cur_pos] = y;
		 }
		else if(i_markmode == 1)
		{
			x_s[0]= x;
			y_s[0]= y;
			i_markmode = 0;
			b_null = true;
			SetClassLong(hvpanel,GCL_HCURSOR,(LONG) hcross); 
			UpdateAviInfo();
		}
		else if(i_markmode == 2)
		{
			x_s[1] = x;
			y_s[1] = y;
			i_markmode = 3;
			SetClassLong(hvpanel,GCL_HCURSOR,(LONG) i_cross2); 
		}	
		else if(i_markmode == 3)
		{
			x_s[2] = x;
			y_s[2] = y;
			i_markmode = 0;
			b_cali = true;
			SetClassLong(hvpanel,GCL_HCURSOR,(LONG)hcross); 
			UpdateAviInfo();
		}	
		else if(i_markmode == 4)
		{
			x_s[3] = x;
			y_s[3] = y;
		    i_markmode = 5;
			SetClassLong(hvpanel,GCL_HCURSOR,(LONG) i_cross2); 
		}
		else if(i_markmode == 5)
		{
			x_s[3] = x-x_s[3];
			y_s[3] = y-y_s[3];	
		    //Normieren
			float f_length;
			f_length = sqrt(pow(x_s[3],2)+pow(y_s[3],2));
			f_alpha = acos((-1)*((float)x_s[3])/f_length);
			//Drehrichtung 
			if (y_s[3] > 0)
			{
			 f_alpha *= (-1);
			}
			if (f_alpha < -M_PI/2)
			{
			 f_alpha += M_PI;
			}
			if (f_alpha > M_PI/2)
			{
			 f_alpha -= M_PI;
			}			
			char cmd[30];
			sprintf(cmd,"Das Koordinatensystem wird um %.2f Grad\ngegen den Uhrzeigersinn gedreht!",f_alpha/M_PI*180);
			MessageBox(NULL,cmd,"",MB_ICONINFORMATION);
			i_markmode = 0;
			SetClassLong(hvpanel,GCL_HCURSOR,(LONG)hcross); 
		}				
		return 0;
	}
    case WM_RBUTTONUP:
	{
	
		if (b_already)
		{
			int n_pos;
			n_pos = SendMessage(c_slide, TBM_GETPOS, 0, 0);
			SendMessage(c_slide, TBM_SETPOS, true, n_pos+1);	
			SendMessage(hmain, WM_HSCROLL,NULL,NULL);					
		}
	 /*if (b_already && (info_avi.cur_pos < info_avi.i_frames-1))
				stepAVI(1,false);
					if (b_dfunc)
					{
				      SendMessage(hmain,WM_PAINT,0,0);		
					}*/
	  return 0;		
	}   
   }
  }
 return DefWindowProc(hWnd, umsg, wParam, lParam);

}
LRESULT CALLBACK WndProc(HWND hWnd, UINT umsg, WPARAM wParam, LPARAM lParam)
{
   HMENU hMenu = GetMenu(hWnd);
   UINT uState = GetMenuState(hMenu,ID_ZOOM,MF_BYCOMMAND);
   
   switch (umsg)
   {
      case WM_PAINT:
	    if(b_already)
		{
	     stepAVI(0,false);
				if (b_dfunc)
				{
				 DrawFunction();
				}
		}
		break;     

     case WM_COMMAND :
	 {
		
         switch( LOWORD( wParam ) )
           {
			case BACK:
			{
				/*if (b_already && (info_avi.cur_pos > 0))
				 stepAVI(-1,false);
				if (b_dfunc)
				{
				    SendMessage(hmain,WM_PAINT,0,0);		
				}*/
				if (b_already)
				{
				 int n_pos;
				 n_pos = SendMessage(c_slide, TBM_GETPOS, 0, 0);
				 SendMessage(c_slide, TBM_SETPOS, true, n_pos-1);	
				 SendMessage(hmain, WM_HSCROLL,NULL,NULL);					
				}
				return 0;
		    }
			case FOR:
			{
			   /*if (b_already && (info_avi.cur_pos < info_avi.i_frames-1))
				stepAVI(1,false);
				if (b_dfunc)
				{
				  SendMessage(hmain,WM_PAINT,0,0);		
				}	*/
				if (b_already)
				{
				 int n_pos;
				 n_pos = SendMessage(c_slide, TBM_GETPOS, 0, 0);
				 SendMessage(c_slide, TBM_SETPOS, true, n_pos+1);	
				 SendMessage(hmain, WM_HSCROLL,NULL,NULL);					
				}
				return 0;
			}
			case ZERO:
			{
			    SetClassLong(hvpanel,GCL_HCURSOR,(LONG) i_null); 
				i_markmode = 1;
				return 0;
			}
			case ZEROT:
			{
			    i_tnull = info_avi.cur_pos;
				UpdateAviInfo();
				return 0;
			}			
			case CALI:
			{
				SetClassLong(hvpanel,GCL_HCURSOR,(LONG) i_cross1); 
				i_markmode = 2;
				return 0;
			}
			case DRAW:
			{
			    ShowWindow(t_function,SW_HIDE);
				ShowWindow(t_start,SW_HIDE);
				ShowWindow(t_end,SW_HIDE);
				ShowWindow(l_x,SW_HIDE);
				ShowWindow(cmdDrawF,SW_HIDE);
				ShowWindow(hinfo,SW_SHOW);
				GetPara();
				DrawFunction();
				b_dfunc = true;
				return 0;
			}
			case ID_FILE_OPEN:
			{
			  OpenFileBox(hWnd);
			  return 0;
			}
			case ID_COORD_ROTATE:
			{
			  SetClassLong(hvpanel,GCL_HCURSOR,(LONG) i_cross1); 
			  i_markmode = 4;
			  return 0;
			}
			case ID_COORD_BACK:
			{
			  if (b_already)
			  {
			   x_values[info_avi.cur_pos] = 0;
		       y_values[info_avi.cur_pos] = 0;
			  }
			  return 0;
			}				
			case ID_PARABEL:
			{
			    ShowWindow(t_function,SW_SHOW);
				ShowWindow(cmdDrawF,SW_SHOW);
				ShowWindow(l_x,SW_SHOW);
				ShowWindow(t_start,SW_SHOW);
				ShowWindow(t_end,SW_SHOW);
				ShowWindow(hinfo,SW_HIDE);
			  return 0;
			}	
			case ID_ZOOM:
			{
			  
			  if (!(uState & MF_CHECKED)) 
			  {
			   CheckMenuItem(hMenu, (UINT) ID_ZOOM, MF_BYCOMMAND | MF_CHECKED); 
			   b_zoom = true;
			  }
			  else
			  {
			   CheckMenuItem(hMenu, (UINT) ID_ZOOM, MF_BYCOMMAND | MF_UNCHECKED); 
			   b_zoom = false;
			  }
			  setZoom(true);
              return 0;
			}
			case ID_FILE_EXPORT:
			{
			  ExportData(false);
			  return 0;
			}
			case ID_FILE_XML:
			{
			  ExportData(true);
			  return 0;
			}
			case ID_FILE_EXIT:
			{
			  int i_answer;
			  i_answer = MessageBox(NULL,"Sind Sie sicher, dass KineVideo geschlossen werden soll?","Daten gespeichert?",MB_YESNO|MB_ICONQUESTION);
			  if (i_answer == IDYES)
			  {
			   SendMessage(hmain,WM_DESTROY,0,0);
			  }
			  return 0;
			}

		 } 
	  }
   case WM_HSCROLL :

      int n_pos;
	  n_pos = SendMessage(c_slide, TBM_GETPOS, 0, 0);
	  stepAVI(n_pos,true);
	  if (b_dfunc)
	  {
		SendMessage(hmain,WM_PAINT,0,0);		
	   }	
      return 0;
 
   case WM_DESTROY:
      {
	    char cmd[300];
        wsprintf(cmd, "close myAVI");
        mciCmd(cmd);
     	PostQuitMessage(0);
         return 0;
      }
   }
   return DefWindowProc(hWnd, umsg, wParam, lParam);
}
HWND CreateToolTip(HWND toolID, HWND hDlg, PTSTR pszText,HINSTANCE hInstance)
{
    if (!toolID || !hDlg || !pszText)
    {
        return FALSE;
    }

    // Create the tooltip. hInstance is the global instance handle.
    HWND hwndTip = CreateWindowEx(NULL, TOOLTIPS_CLASS, NULL,
                              WS_POPUP |TTS_ALWAYSTIP | TTS_BALLOON,
                              CW_USEDEFAULT, CW_USEDEFAULT,
                              CW_USEDEFAULT, CW_USEDEFAULT,
                              hDlg, NULL, 
                              hInstance, NULL);
    
    
    // Associate the tooltip with the tool.
    TOOLINFO toolInfo = { 0 };
    toolInfo.cbSize = sizeof(toolInfo);
    toolInfo.hwnd = hDlg;
    toolInfo.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
    toolInfo.uId = (UINT_PTR)toolID;
    toolInfo.lpszText = pszText;
    SendMessage(hwndTip, TTM_ADDTOOL, 0, (LPARAM)&toolInfo);

    return hwndTip;
}
