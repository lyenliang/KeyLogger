#include "functions.h"

using namespace std;

string intToString(int i){
	char buffer[4];
	_itoa_s(i, buffer, 10);
	return string(buffer);
}

string getCurrDir(){
	char *curdir = new char[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, curdir);
	string rv(curdir);
	delete[] curdir;
	return rv;
}

string getSelfPath(){
	char selfpath[MAX_PATH];
	GetModuleFileName(NULL, selfpath, MAX_PATH);
	return string(selfpath);
}

string dirBasename(string path){
	if(path.empty())
		return string("");
	
	if(path.find("\\") == string::npos)
		return path;
	
	if(path.substr(path.length() - 1) == "\\")
		path = path.substr(0, path.length() - 1);
	
	size_t pos = path.find_last_of("\\");
	if(pos != string::npos)
		path = path.substr(0, pos);
	
	if(path.substr(path.length() - 1) == "\\")
		path = path.substr(0, path.length() - 1);
	
	return path;
}

bool isCapsLock() {
	return (GetKeyState(VK_CAPITAL) & 0x0001) != 0;  // If the low-order bit is 1, the key is toggled
}

bool isShift() {
	return (GetKeyState(VK_SHIFT) & 0x8000) != 0; // If the high-order bit is 1, the key is down; otherwise, it is up.
}

void logFile(ofstream& outFile, string msg) {
		outFile << msg;
	#ifdef DEBUG
		cout << msg;
	#endif
}


