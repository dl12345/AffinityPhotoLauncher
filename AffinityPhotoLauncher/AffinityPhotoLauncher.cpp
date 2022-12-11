#include <ShObjIdl.h>
#include <atlbase.h>
#include <appmodel.h>
#include <string>
#include <iostream>

class CommandLineArgs
{
protected:

    std::wstring m_Commandline;
    std::wstring m_Fullname;
    std::wstring m_Basename;
    std::wstring m_Appname;
    std::wstring m_Args;

public:

    CommandLineArgs();
    ~CommandLineArgs() {}

    inline const std::wstring Fullname() const { return m_Fullname; }
    inline const std::wstring Basename() const { return m_Basename; }
    inline const std::wstring Appname() const { return m_Appname; }
    inline const std::wstring Arguments() const { return m_Args; }
};

CommandLineArgs::CommandLineArgs()
{
    // GetCommandLineA returns quoted arguments if there are spaces
    // Check if the first character is a quote then find the first occurrence of '\" '
    // or if unquoted, ' ', to locate the beginning of the arguments

    m_Commandline = GetCommandLineW();
    size_t offset = 0;

    // separate out program name and arguments

    if (m_Commandline[0] == '\"') // quoted program name
    {
        offset = m_Commandline.find(L"\" ", 1);
        if (offset != std::string::npos)
        {
            m_Fullname = m_Commandline.substr(1, offset - 1);
            m_Args = m_Commandline.substr(offset + 2);
        }

    }
    else // unquoted
    {
        offset = m_Commandline.find(L" ", 1);
        if (offset != std::string::npos)
        {
            m_Fullname = m_Commandline.substr(0, offset);
            m_Args = m_Commandline.substr(offset + 1);
        }

    }

    // get the basename of the program 

    offset = m_Fullname.find_last_of(L'\\');
    if (offset != std::string::npos)
    {
        m_Basename = m_Fullname.substr(offset + 1);
    }
    else 
    {
        m_Basename = m_Fullname;
    }

    // strip off the file extension

    offset = m_Basename.find_last_of(L".exe");
    if (offset != std::string::npos)
    {
        m_Appname = m_Basename.substr(0, offset - 3);
    }
    else
    {
        m_Appname = m_Basename;
    }

#ifdef DEBUG
    std::wcout << L"Fullname: " << m_Fullname << std::endl;
    std::wcout << L"Basename: " << m_Basename << std::endl;
    std::wcout << L"Appname: " << m_Appname << std::endl;
    std::wcout << L"Commandline: " << m_Commandline << std::endl;
    std::wcout << L"Args: " << m_Args << std::endl;
#endif

}


class UWPAppLauncher
{
protected:
    
    std::wstring    m_Appname;
    std::wstring    m_PublisherName;
    std::wstring    m_KeyName;
    std::wstring    m_Fullname;
    HRESULT         m_hr;
    bool            m_Installed;

protected: 

    const bool PackageFullNameFromFamilyName();
    inline const bool Success() const { return (SUCCEEDED(m_hr) ? true : false); }

public:
   
    UWPAppLauncher(const std::wstring appName, const std::wstring publisherName, const std::wstring keyName);
    ~UWPAppLauncher() {} 

    inline const std::wstring AppName() const { return m_Appname; }
    inline const std::wstring PublisherName() const { return m_PublisherName; }
    inline const std::wstring KeyName() const { return m_KeyName; }
    inline const std::wstring PackageFamilyName() const;
    inline const std::wstring PackageFullName() const;
    inline const std::wstring AppAMUID() const;
    inline const bool IsInstalled() const;
    const bool Launch(const std::wstring args);
    const std::wstring GetErrorMessage() const; 
};

const bool UWPAppLauncher::PackageFullNameFromFamilyName()
{
    UINT32 count = 0;
    UINT32 length = 0;

    m_hr = GetPackagesByPackageFamily(PackageFamilyName().c_str(), &count, nullptr, &length, nullptr);

    if (SUCCEEDED(m_hr))
    {

        PWSTR* fullNames = (PWSTR*)malloc(count * sizeof(*fullNames));
        PWSTR buffer = (PWSTR)malloc(length * sizeof(WCHAR));
        UINT32* properties = (UINT32*)malloc(count * sizeof(*properties));

        if (fullNames && buffer && properties)
        {
            m_hr = GetPackagesByPackageFamily(PackageFamilyName().c_str(), &count, fullNames, &length, buffer);

            if (SUCCEEDED(m_hr) && count > 0)
            {
                m_Fullname = std::wstring(fullNames[0]);
            }
        }

        if (properties != nullptr) free(properties);

        if (buffer != nullptr) free(buffer);

        if (fullNames != nullptr) free(fullNames);

    }
    return SUCCEEDED(m_hr) && count > 0 ? true : false;
}

UWPAppLauncher::UWPAppLauncher(const std::wstring appName, const std::wstring publisherName, const std::wstring keyName)
    :   m_Appname(appName),
        m_PublisherName(publisherName),
        m_KeyName(keyName),
        m_Installed(false)
{
    m_Installed = PackageFullNameFromFamilyName();
}

inline const std::wstring UWPAppLauncher::PackageFamilyName() const
{
    return m_PublisherName + L"." + m_Appname + L"_" + m_KeyName;
}

inline const std::wstring UWPAppLauncher::AppAMUID() const
{
    return PackageFamilyName() + L"!" + m_PublisherName + L"." + m_Appname;
}

inline const bool UWPAppLauncher::IsInstalled() const
{
    return m_Installed;
}

const bool UWPAppLauncher::Launch(const std::wstring args)
{
    m_hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    bool result = false;

    if (SUCCEEDED(m_hr))
    {
        CComPtr<IApplicationActivationManager> AppActivationMgr = nullptr;

        if (SUCCEEDED(m_hr))
        {
            m_hr = CoCreateInstance(
                CLSID_ApplicationActivationManager,
                nullptr,
                CLSCTX_LOCAL_SERVER,
                IID_PPV_ARGS(&AppActivationMgr)
            );
        }
        if (SUCCEEDED(m_hr))
        {
            DWORD pid = 0;
            m_hr = AppActivationMgr->ActivateApplication(AppAMUID().c_str(), args.c_str(), AO_NONE, &pid);
        }
    }

    CoUninitialize();
    if (SUCCEEDED(m_hr)) result = true;
    return result;
}

const std::wstring UWPAppLauncher::GetErrorMessage() const
{
    static WCHAR buf[512];
    buf[0] = '\0';

    DWORD cchMsg = FormatMessage(

        FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL,
        m_hr,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        buf,
        512,
        NULL
    );

    return std::wstring(buf);
}

int wmain(int argc, WCHAR *argv[])
{
    int result = 1;
    CommandLineArgs args;
   
    std::wstring retail = L"3cqzy0nppv2rt";
    std::wstring msstore = L"844sdzfcmm7k0";
    
    UWPAppLauncher retailapp(args.Appname(), L"SerifEuropeLtd", retail);
    UWPAppLauncher msstoreapp(args.Appname(), L"SerifEuropeLtd", msstore);

    UWPAppLauncher* p = nullptr;

    retailapp.IsInstalled() ? p = &retailapp : (msstoreapp.IsInstalled() ? p = &msstoreapp : p = nullptr);

    if (p == nullptr)
    {
        std::wcout << args.Appname() << L" is not installed\n";
        std::wcout << L"Ensure the binary is named one of AffinityPhoto2.exe, AffinityDesigner2.exe or AffinityPublisher2.exe depending on which application you wish to launch" << std::endl;
    }
    else if (!p->Launch(args.Arguments()))
    {
        std::wcout << L"Failed to launch app: " << p->GetErrorMessage() << std::endl;
    }
    else result = 0;

    return result;
}


