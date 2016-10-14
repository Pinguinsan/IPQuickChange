#ifndef CHANGEIPSETTINGS_H
#define CHANGEIPSETTINGS_H

#include <string>

short __stdcall executeStaticChange(char *IP, char *subnet, char *domain);
short __stdcall executeDynamicChange();
short __stdcall generateDeviceList(char *IPBase, short IPTail);
short __stdcall parseDeviceListToFile(char *IPBase);
short __stdcall cleanUpTmp();
short __stdcall delay(double howLong);

std::string parseIPFromRawString(const std::string &stringToParse, const std::string &IPBaseArg);
bool isNumeric(const char &charToCheck);

bool directoryExists(const std::string &directoryToCheck);

template<typename T>
bool isNull(T *ptrToCheck);

template<typename T>
std::string toString(const T &input);


#endif