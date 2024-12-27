// dllmain.cpp : 定义 DLL 应用程序的入口点。
#include "pch.h"
#include <atlbase.h>
#include <string>
#include <windows.h>  // For MessageBoxW
#include <shellapi.h>      // 引入 shellapi.h 头文件
#include <shlobj.h>
#include <sstream>

#import "vgcoreauto.tlb"
long majorVersion, minorVersion;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

class CVGAppPlugin : public VGCore::IVGAppPlugin
{
private:
    VGCore::IVGApplication * m_pApp;
    ULONG m_ulRefCount;
    long m_lCookie;
    bool m_bEnabled;

    bool CheckSelection();
    void OnSmartLaser();
    void ExportSelectionAsDXF();
    std::wstring getInstallPathFromRegistry();
    std::wstring GetDocumentsFolderPath();

public:
    CVGAppPlugin();

    // IUnknown
public:
    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObject);
    STDMETHOD_(ULONG, AddRef)(void) { return ++m_ulRefCount; }
    STDMETHOD_(ULONG, Release)(void)
    {
        ULONG ulCount = --m_ulRefCount;
        if (ulCount == 0)
        {
            delete this;
        }
        return ulCount;
    }

    // IDispatch
public:
    STDMETHOD(GetTypeInfoCount)(UINT* pctinfo) { return E_NOTIMPL; }
    STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) { return E_NOTIMPL; }
    STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) { return E_NOTIMPL; }
    STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);
    BOOL IsInstalledInProgramFiles();
    // IVGAppPlugin
public:
    STDMETHOD(raw_OnLoad)(VGCore::IVGApplication* Application);
    STDMETHOD(raw_StartSession)();
    STDMETHOD(raw_StopSession)();
    STDMETHOD(raw_OnUnload)();
};

bool CVGAppPlugin::CheckSelection()
{
    bool bRet = false;
    if (m_pApp->Documents->Count > 0)
    {
        bRet = (m_pApp->ActiveSelection->Shapes->Count > 0);
    }
    return bRet;
}

std::wstring CVGAppPlugin::GetDocumentsFolderPath()
{
    PWSTR path = NULL;
    HRESULT hr = SHGetKnownFolderPath(FOLDERID_Documents, 0, NULL, &path);
    if (SUCCEEDED(hr))
    {
        std::wstring documentsPath(path);
        CoTaskMemFree(path);  // 释放路径
        return documentsPath;
    }
    else
    {
        return L"";
    }
}

std::wstring CVGAppPlugin::getInstallPathFromRegistry()
{
    HKEY hKey;
    std::wstring installPath;
    // 注册表路径
    const wchar_t* regPath = L"SOFTWARE\\WOW6432Node\\SmartLaser";
    const wchar_t* valueName = L"InstallPath";

    // 打开注册表键
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, regPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        wchar_t pathBuffer[MAX_PATH];
        DWORD bufferSize = sizeof(pathBuffer);
        DWORD valueType;

        // 查询键值
        if (RegQueryValueExW(hKey, valueName, nullptr, &valueType, reinterpret_cast<LPBYTE>(pathBuffer), &bufferSize) == ERROR_SUCCESS)
        {
            if (valueType == REG_SZ) // 确保是字符串类型
            {
                installPath = pathBuffer;
            }
        }

        // 关闭注册表键
        RegCloseKey(hKey);
    }

    return installPath;
}

void CVGAppPlugin::ExportSelectionAsDXF()
{
    try
    {
        // 获取当前文档对象（假设文档已打开）
        VGCore::IVGDocumentPtr pDoc = m_pApp->ActiveDocument;


        HRESULT hr = CoInitialize(NULL);  // 或者 CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
        if (FAILED(hr))
        {
            MessageBox(NULL, "Failed to initialize COM", "Error", MB_ICONSTOP);
            return;
        }
        pDoc->PutUnit(VGCore::cdrUnit::cdrMillimeter);

        // 使用 CreateStructExportOptions 创建导出选项
        VGCore::IVGStructExportOptionsPtr pExportOptions = m_pApp->CreateStructExportOptions();
        if (pExportOptions == NULL)
        {
            MessageBox(NULL, "Failed to create instance of IVGStructExportOptions", "Error", MB_ICONSTOP);
            CoUninitialize();
            return;
        }

        // 设置导出范围为选中区域（根据需要调整）
        VGCore::cdrExportRange exportRange = VGCore::cdrSelection;  // 如果想导出全部内容可以使用 cdrExportRangeAll


        // 使用 CreateStructPaletteOptions 创建调色板选项
        VGCore::IVGStructPaletteOptionsPtr pPaletteOptions = m_pApp->CreateStructPaletteOptions();
        if (pPaletteOptions == NULL)
        {
            MessageBox(NULL, "Failed to create instance of IVGStructPaletteOptions", "Error", MB_ICONSTOP);
            CoUninitialize();
            return;
        }
       
        // 设置导出格式为 DXF
        VGCore::cdrFilter exportFilter = VGCore::cdrDXF;  // 设置格式为 DXF

        // 设置导出文件路径
        std::wstring dxfFilePath = GetDocumentsFolderPath() + L"\\document_output.dxf";

        // 设置其他导出选项（根据需求设置）
        pExportOptions->PutOverwrite(VARIANT_TRUE);  // 如果文件已存在则覆盖

        // 导出文件
        hr = pDoc->Export(dxfFilePath.c_str(), exportFilter, exportRange, pExportOptions, pPaletteOptions);
        if (FAILED(hr))
        {
            MessageBox(NULL, "Failed to export as DXF", "Error", MB_ICONSTOP);
            CoUninitialize();
            return;
        }
        std::wstring installPath = getInstallPathFromRegistry();
        std::wstring exePath = installPath + L"\\SmartLaser.exe";
        ShellExecuteW(NULL, L"open", exePath.c_str(), dxfFilePath.c_str(), NULL, SW_SHOWNORMAL);
      
        // 清理 COM 初始化
        CoUninitialize();
    }
    catch (_com_error& e)
    {
        MessageBox(NULL, e.Description(), "错误", MB_ICONSTOP);
    }
}


void CVGAppPlugin::OnSmartLaser()
{
}

CVGAppPlugin::CVGAppPlugin() :
    m_pApp(NULL),
    m_lCookie(0),
    m_ulRefCount(1),
    m_bEnabled(false)
{
}

STDMETHODIMP CVGAppPlugin::QueryInterface(REFIID riid, void** ppvObject)
{
    HRESULT hr = S_OK;
    m_ulRefCount++;
    if (riid == IID_IUnknown)
    {
        *ppvObject = (IUnknown*)this;
    }
    else if (riid == IID_IDispatch)
    {
        *ppvObject = (IDispatch*)this;
    }
    else if (riid == __uuidof(VGCore::IVGAppPlugin))
    {
        *ppvObject = (VGCore::IVGAppPlugin*)this;
    }
    else
    {
        m_ulRefCount--;
        hr = E_NOINTERFACE;
    }
    return hr;
}

STDMETHODIMP CVGAppPlugin::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    switch (dispIdMember)
    {
    case 0x0011: //DISPID_APP_SELCHANGE
        m_bEnabled = CheckSelection();
        break;

    case 0x0014: // DISPID_APP_ONPLUGINCMD
        if (pDispParams != NULL && pDispParams->cArgs == 1)
        {
           
            _bstr_t strCmd(pDispParams->rgvarg[0].bstrVal);
            if (strCmd == _bstr_t("SmartLaser"))
            {        
                ExportSelectionAsDXF();     
            }
        }
        break;

    case 0x0015: // DISPID_APP_ONPLUGINCMDSTATE
        if (pDispParams != NULL && pDispParams->cArgs == 3)
        {
            _bstr_t strCmd(pDispParams->rgvarg[2].bstrVal);
            if (strCmd == _bstr_t("SmartLaser"))
            {
                *pDispParams->rgvarg[1].pboolVal = m_bEnabled ? VARIANT_TRUE : VARIANT_FALSE;
            }
        }
        break;

    case 0x0017: // DISPID_APP_ONPLUGINCMDSTATE
        MessageBox(NULL, "On Click", ("Error"), MB_ICONSTOP);
        break;
    }
    return S_OK;
}

void ShowCorelDrawVersion(long majorVersion, long minorVersion)
{
    std::wstring versionMessage;

    if (majorVersion == 14)
    {
        versionMessage = L"CorelDRAW X4 (Version: " + std::to_wstring(majorVersion) + L"." + std::to_wstring(minorVersion) + L")";
    }
    else if (majorVersion == 15)
    {
        versionMessage = L"CorelDRAW X5 (Version: " + std::to_wstring(majorVersion) + L"." + std::to_wstring(minorVersion) + L")";
    }
    else
    {
        versionMessage = L"CorelDRAW Other Version (Version: " + std::to_wstring(majorVersion) + L"." + std::to_wstring(minorVersion) + L")";
    }
    // 使用 MessageBox 显示版本信息
    MessageBoxW(NULL, versionMessage.c_str(), L"CorelDRAW Version", MB_OK | MB_ICONINFORMATION);
}

STDMETHODIMP CVGAppPlugin::raw_OnLoad(VGCore::IVGApplication* Application)
{
    m_pApp = Application;
    if (m_pApp)
    {
        m_pApp->AddRef();

        // 获取 CorelDRAW 版本
        long majorVersion, minorVersion;
        m_pApp->get_VersionMajor(&majorVersion);
        m_pApp->get_VersionMinor(&minorVersion);

//        ShowCorelDrawVersion(majorVersion, minorVersion);

        // 根据版本适配接口
        if (majorVersion >= 20)
        {
            // 使用新版本的功能
        }
        else
        {
            // 使用旧版本的功能
        }
    }
    return S_OK;
}
VGCore::ICUICommandBarPtr newToolbar;
STDMETHODIMP CVGAppPlugin::raw_StartSession()
{
    try
    {
        // 检查是否已经存在一个名为 "SmartLaserToolbar" 的工具栏
        bool toolbarExists = false;
        

        for (long i = 1; i <= m_pApp->CommandBars->Count; ++i)
        {
            auto toolbar = m_pApp->CommandBars->Item[i];
            if (toolbar->GetName() == _bstr_t("SmartLaserToolbar"))
            {
                toolbarExists = true;
                newToolbar = toolbar;
                break;
            }
        }


        if (!toolbarExists)
        {
            newToolbar = m_pApp->CommandBars->Add(_bstr_t("SmartLaserToolbar"), VGCore::cuiBarLeft, VARIANT_TRUE);
            newToolbar->Visible = VARIANT_TRUE; // 显示工具栏
        }

        bool buttonExists = false;
        for (long i = 1; i <= newToolbar->Controls->Count; ++i)
        {
            auto ctl = newToolbar->Controls->Item[i];
            if (ctl->GetCaption() == _bstr_t("SmartLaser"))
            {
                buttonExists = true;
                break;
            }
        }

        {
            m_pApp->AddPluginCommand(_bstr_t("SmartLaser"), _bstr_t("SmartLaser"), _bstr_t("Launch Smart Laser"));
            
            // 在新工具栏中添加按钮
            VGCore::ICUIControlPtr ctl;
            if (majorVersion == 14 || majorVersion == 15)
            {
                ctl = newToolbar->Controls->AddCustomButton(
                    VGCore::cdrCmdCategoryPlugins,
                    _bstr_t("SmartLaser"),
                    1, // 按钮 ID 或图标索引，确保唯一
                    VARIANT_TRUE
                );
            }
            else
            {
                ctl = newToolbar->Controls->AddCustomButton(
                    VGCore::cdrCmdCategoryPlugins,
                    _bstr_t("SmartLaser"),
                    1, // 按钮 ID 或图标索引，确保唯一
                    VARIANT_FALSE
                );
            }
//            MessageBox(NULL, "4", "Error", MB_ICONERROR);

            // 设置自定义图标
            _bstr_t bstrPathX;
            #ifdef _WIN64
                bstrPathX = (m_pApp->Path + _bstr_t("Plugins64\\SLRicon.bmp"));  // 使用正确的图标路径
                HRESULT hr = ctl->SetIcon2(bstrPathX);
            #else
                bstrPathX = (m_pApp->Path + _bstr_t("Plugins\\SLRicon.bmp"));  // 使用正确的图标路径
                HRESULT hr = ctl->SetCustomIcon(bstrPathX);
            #endif
                    
 //               HRESULT hr = ctl->SetCustomIcon(bstrPathX);
//                MessageBox(NULL, "5", "Error", MB_ICONERROR);
 //                            ctl->SetCustomIcon(bstrPathX86);
                if (FAILED(hr))
                {
                    MessageBox(NULL, "Failed to set icon!", "Error", MB_ICONERROR);
                }
    
            //else
            //{
            //    MessageBox(NULL, "Icon file not found!", "Error", MB_ICONERROR);
            //}
        }

        // 订阅事件
        m_lCookie = m_pApp->AdviseEvents(this);
    }
    catch (_com_error& e)
    {
        MessageBox(NULL, e.Description(), "Error", MB_ICONSTOP);
    }
    return S_OK;
}


STDMETHODIMP CVGAppPlugin::raw_StopSession()
{
    try
    {
        // 取消事件订阅
        m_pApp->UnadviseEvents(m_lCookie);

        // 移除插件命令
        m_pApp->RemovePluginCommand(_bstr_t("SmartLaser"));
        //// 清理工具栏和按钮
        //for (long i = 1; i <= m_pApp->CommandBars->Count; ++i)
        //{
        //    auto toolbar = m_pApp->CommandBars->Item[i];
        //    if (toolbar->GetName() == _bstr_t("SmartLaserToolbar"))
        //    {
        //        toolbar->Delete();  // 删除工具栏
        //        break;
        //    }
        //}
    }
    catch (_com_error& e)
    {
        MessageBox(NULL, e.Description(), _T("Error"), MB_ICONSTOP);
    }
    return S_OK;
}

STDMETHODIMP CVGAppPlugin::raw_OnUnload()
{
    if (m_pApp)
    {
        m_pApp->Release();
        m_pApp = NULL;
    }
    return S_OK;
}

extern "C" __declspec(dllexport) DWORD APIENTRY AttachPlugin(VGCore::IVGAppPlugin * *ppIPlugin)
{
    *ppIPlugin = new CVGAppPlugin;
    return 0x100;
}