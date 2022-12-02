#include <iostream>
#include <fstream>
#include <sstream>
#include <cstring>
#include <string>
#include <shlobj.h>
#include <windows.h>
#include <direct.h>
#include <urlmon.h>
#include <ctime>
using namespace std;

ofstream f_log;

const int F_SZ = 256, PROD_SZ = 30, INFO_SZ = 512 /*32767*/; /// TODO: sizeof(arr) / sizeof(arr[0])

void pause() {
    cout << "\nPress Enter to continue.\n";
    cin.get();
}

string header(string name, bool first=false) {
    ostringstream out;
    string decor (25, '=');
    if(!first) out << '\n';
    out << decor << ' ' << name << ' ' << decor << '\n';
    return out.str();
}

ostringstream Debug;
/// close and
void writeDebug(const char *err=NULL) {
    if (!f_log.is_open()) return;
    f_log << header("Debug");
    f_log << Debug.str();
    if (err != NULL)
        f_log << string(10, '~') << ' ' << "Errors" << ' ' << string(20, '~') << '\n' << err;
    f_log.close();
}
void stop(string err) {
    writeDebug(err.c_str());
    cout << err;
    pause();
    exit(1);
}

BOOL WINAPI onExit(DWORD s) {
    writeDebug(); // and close
    if (s == CTRL_CLOSE_EVENT || s == CTRL_LOGOFF_EVENT || s == CTRL_SHUTDOWN_EVENT) return FALSE;
    SetWindowPos(GetConsoleWindow(), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    //exit(1);
    return FALSE;
}

bool oFile(ifstream & ifs, string name) {
    ifs.open(name);
    return ifs.is_open();
}

/*void readLines(string buf, void(* cb)(string)) {
    string::size_type s1, s2 = 0;
    while (true) {
        s1 = s2;
        s2 = buf.find('\n', s1);
        if (s2 == string::npos) break;
        cb(buf.substr(s1, ++s2-s1-1));
    }
    cb(buf.substr(s1));
}*/

char Fname[F_SZ], InfoBuf[INFO_SZ];
string LogfilePrefix = "activation_";
string UNKNOWN = "Unknown";

const int VAR_C = 3, SKIP_SZ = 6; // CONST_
char Vars[][INFO_SZ] = {"MSG_ACTSUCCESS", "MSG_OFFLINEACTSUCCESS", "MSG_KEYINSTALLSUCCESS"};
char Vals[VAR_C][INFO_SZ];
int ParsedVars = 0;

string TITLE = "Microsoft Office activation wizard 2.0";
string OFFICE16x86 = "C:\\Program Files (x86)\\Microsoft Office\\Office16\\";
string OFFICE16 = "C:\\Program Files\\Microsoft Office\\Office16\\";
string OSPP = "OSPP.VBS";
string DINSTID_SEP = "edition: ";
string DCONFID_SEP = "\"confirmation_id_no_dash\":\"";

void parseVar(char * buf) {
    buf += SKIP_SZ;
    int sz = strlen(Vars[ParsedVars]);
    if (strncmp(Vars[ParsedVars], buf, sz) == 0) {
        buf += sz;
        char * pos1 = strstr(buf, "= \""); // strlen = 3
        char * pos2 = strrchr(buf, '"');
        if (pos1 == NULL || pos2 == NULL || pos1 >= pos2) {
            stop("Error parsing the " + OSPP + " file");
        }
        pos1[pos2-pos1] = 0;
        //Vals[ParsedVars++] = pos1 + 3; // +3 from strlen
        strncpy(Vals[ParsedVars++], pos1+3, pos2-pos1);
    }
}

void printD(const char* h) {
    Debug << string(3, '+') << ' ' << h << ' ' << string(10, '+') << '\n';
}
void endD(const char* h, string& result) {
    Debug << result << '\n' << string(15 + strlen(h), '+') << '\n';
}

// source: https://stackoverflow.com/a/51959694
string req(const char* url) {
    printD(url);
    IStream* stream;
    HRESULT ok = URLOpenBlockingStreamA(0, url, &stream, 0, 0);
    if (ok != 0)
        stop("Error sending http request");
    char buffer[128];
    unsigned long bytesRead;
    ostringstream result;
    while (stream->Read(buffer, sizeof(buffer), &bytesRead) == 0 && bytesRead > 0U) {
        result.write(buffer, bytesRead);
    }
    stream->Release();
    string strR = result.str();
    endD(url, strR);
    return strR;
}

// source: https://stackoverflow.com/a/478960
string exec(const char* cmd) {
    printD(cmd);
    char buffer[128];
    ostringstream result;
    FILE* pipe = popen(cmd, "r");
    if (!pipe) /*throw runtime_error*/ stop("popen() failed!");
    try {
        while (fgets(buffer, sizeof(buffer), pipe) != NULL) {
            result << buffer;
        }
    } catch (...) {
        pclose(pipe);
        stop("Broken pipe");
        //throw;
    }
    pclose(pipe);
    string strR = result.str();
    endD(cmd, strR);
    return strR;
}

int main(int argc, char **args) {
    //ShowWindow(GetConsoleWindow(), SW_HIDE);
    GetModuleFileNameA(NULL, Fname, F_SZ);
    SetConsoleTitleA(TITLE.c_str());
    if(IsUserAnAdmin()){
        SetWindowPos(GetConsoleWindow(), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        SetConsoleCtrlHandler(onExit, TRUE);
        atexit([] {onExit(CTRL_C_EVENT);});
        system("cls");
        cout << TITLE << '\n';

        long unsigned int Info_SZ = INFO_SZ;
        if (!GetComputerNameA(InfoBuf, &Info_SZ)) strncpy(InfoBuf, UNKNOWN.c_str(), INFO_SZ-1);
        f_log.open(LogfilePrefix + InfoBuf + ".log");
        f_log << header("Informations", true);
        f_log << "Computer name: " << InfoBuf << '\n';
        if (!GetUserNameA(InfoBuf, &Info_SZ)) strncpy(InfoBuf, UNKNOWN.c_str(), INFO_SZ-1);
        f_log << "User name: " << InfoBuf << '\n';
        f_log << "Activation wizard path: " << Fname << '\n';

        f_log << OSPP << " path: ";
        ifstream f_ospp;
        if (chdir(OFFICE16x86.c_str()) == -1 || !oFile(f_ospp, OSPP))
            if (chdir(OFFICE16.c_str()) == -1 || !oFile(f_ospp, OSPP))
                stop(OSPP + " not found");
        f_log << getcwd(Fname, F_SZ) << '\\' << OSPP << '\n';

        while (ParsedVars < VAR_C || f_ospp.eof()) {
            f_ospp.getline(InfoBuf, INFO_SZ);
            parseVar(InfoBuf);
        }
        for(int i=0; i<VAR_C; ++i) f_log << Vars[i] << ": " << Vals[i] << '\n';

        time_t now = time(0);
        tm *gmtm = gmtime(&now);
        char* dt = asctime(gmtm);
        f_log << "Date and time: " << dt;

        f_ospp.close();
        char ProdKey[PROD_SZ];
        if (argc > 1) {
            strncpy(ProdKey, args[1], PROD_SZ); // args
            ProdKey[PROD_SZ - 1] = 0; // Null byte to end
            cout << "Product key (from args[1]): " << ProdKey << '\n';
        }
        else {
            cout << "Enter product key: "; // input from stdin
            cin.getline(ProdKey, PROD_SZ);
        }
        int ln = strlen(ProdKey);
        if (ln == 0)
            stop("Empty product key");
        f_log << header("Activation");
        f_log << "Product key: ";
        for (int i=0; ProdKey[i]; ++i) {
            if (ProdKey[i] >= 'a' && ProdKey[i] <= 'z') ProdKey[i] -= 'a' - 'A';
            f_log << ProdKey[i];
            if (
                ProdKey[i] != '-' &&
                (ProdKey[i] < '0' || ProdKey[i] > '9') &&
                (ProdKey[i] < 'A' || ProdKey[i] > 'Z')
                ) {
                    f_log << '\n';
                    stop("Suspicious characters in the product key");
                }
        }
        f_log << '\n';

        cout << "Installing product key.. " << flush;
        if (exec(("cscript /nologo " + OSPP + " /inpkey:" + ProdKey).c_str()).find(Vals[2]) == string::npos)
            stop("Error installing product key");
        cout << "Done.\n";

        cout << "Reading out installation id.. " << flush;
        string dinstid_raw = exec(("cscript /nologo " + OSPP + " /dinstid").c_str());
        string::size_type iid_start = dinstid_raw.rfind(DINSTID_SEP);
        if (iid_start == string::npos) stop("Empty installation id");
        string dinstid = dinstid_raw.substr(iid_start + DINSTID_SEP.size(), 7*9);
        f_log << "Installation ID: " << dinstid << '\n';
        for (string::iterator it = dinstid.begin(); it < dinstid.end(); ++it)
            if (*it < '0' || *it > '9')
                stop("Suspicious characters in the installation id");
        cout << "Done.\n";

        cout << "Requesting the Confirmation ID.. " << flush;
        string dconfid_raw = req(("https://0xc004c008.com/ajax/cidms_api?iids=" + dinstid + "&username=trogiup24h&password=PHO").c_str());
        string::size_type cid_start = dconfid_raw.rfind(DCONFID_SEP);
        if (cid_start == string::npos) stop("Empty confirmation id");
        string dconfid = dconfid_raw.substr(cid_start + DCONFID_SEP.size(), 6*8);
        f_log << "Confirmation ID: " << dconfid << '\n';
        for (string::iterator it = dconfid.begin(); it < dconfid.end(); ++it)
            if (*it < '0' || *it > '9')
                stop("Suspicious characters in the confirmation id");
        cout << "Done.\n";

        cout << "Entering Confirmation ID.. " << flush;
        if (exec(("cscript /nologo " + OSPP + " /actcid:" + dconfid).c_str()).find(Vals[1]) == string::npos)
            stop("Error entering confirmation id");
        cout << "Done.\n";

        cout << "Activating Office.. " << flush;
        if (exec(("cscript /nologo " + OSPP + " /act").c_str()).find(Vals[0]) == string::npos)
            stop("Error activating office");
        cout << "Done!\n";

        cout << header("PRODUCT ACTIVATED SUCCESSFULLY");
        pause();
    } else {
        if((INT_PTR) ShellExecuteA(NULL, LPSTR("runas"), LPSTR(Fname), LPSTR(argc > 1 ? args[1] : ""), NULL, SW_SHOWNORMAL) < 32) {
            cout << "To activate office need administrator privileges\n";
            pause();
            return 1;
        }
    }
    return 0;
}
