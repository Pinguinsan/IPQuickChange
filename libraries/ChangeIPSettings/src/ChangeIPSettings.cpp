#include <iostream>
#include <fstream>
#include <cstdlib>
#include <sstream>
#include <vector>
#include <ctime>
#include <Windows.h>
#include "../include/dateTimeStamps.h"
#include "../include/ChangeIPSettings.h"

#define INVALID_FILE_ATTRIBUTES 0xffffffff

//Private Declare Function executeStaticChange Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" (ByVal IPAddress As String, ByVal SubnetMask As String, ByVal DomainAddress As String) As Long
short __stdcall executeStaticChange(char *IP, char *subnet, char *domain)
{
	if ((!isNull(IP)) && (!isNull(subnet)) && (!isNull(domain)))
	{
		std::string shellCommand = "netsh int ip set address \"Local Area Connection\" static ";
		shellCommand.append(std::string(IP));
		shellCommand.append(" ");
		shellCommand.append(std::string(subnet));
		shellCommand.append(" ");
		shellCommand.append(std::string(domain));
		shellCommand.append(" 1 >> ../tmp/changeresults.txt");
		return 0;
		return static_cast<short>(system(shellCommand.c_str()));
	}
	return -1;
}

//Private Declare Function executeDynamicChange Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" () As Integer
short __stdcall executeDynamicChange()
{
	std::string shellCommand = "netsh int ip set address name=\"Local Area Connection\" source=dhcp >> \"../tmp/changeresults.txt\" & ipconfig /flushdns";
	short returnVal = system(shellCommand.c_str());
	return returnVal;
}

//Private Declare Function generateDeviceList Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" (ByVal IPBase As String, ByVal IPTail As Integer) As Integer
short __stdcall generateDeviceList(char *IPBase, short IPTail)
{
	std::string stringList = "";
	std::string sysCallString = "";
	std::string tempIPString;
	DeleteFile("../tmp/pingresults.txt");
	DeleteFile("../tmp/tmpping.bat");
	DeleteFile("../tmp/parsedips.txt");
	
	std::ofstream writeToFile;
	writeToFile.open("../tmp/tmpping.bat", std::ofstream::app);
	for (int i = 1; i < 255; i++)
	{
		if (i != IPTail)
		{
			sysCallString.append("ping ");
			sysCallString.append(IPBase);
			sysCallString.append(toString(i));
			sysCallString.append(" -n 1 -w 75 >> ../tmp/pingresults.txt");
			writeToFile << sysCallString << std::endl;
			sysCallString = "";
		}
	}
	writeToFile.close();
	/*return static_cast<short>(*/system("cd .. & cd tmp & tmpping.bat");
	return 0;
}

//Private Declare Function parseDeviceListToFile Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" (ByVal IPBase As String) As Integer
short __stdcall parseDeviceListToFile(char *IPBase)
{
	std::string rawInput;
	std::string tempIPLine;
	std::vector<std::string> goodPingVector; 
	std::ifstream readFromFile;
	readFromFile.open("../tmp/pingresults.txt");
	if (readFromFile.is_open())
	{
		while (std::getline(readFromFile, rawInput))
		{
			if (rawInput.substr(0,7) == "Pinging") //Pinging header, IP comes directly after
			{
				tempIPLine = rawInput;
				std::getline(readFromFile, rawInput);
				std::getline(readFromFile, rawInput);
				if (rawInput.substr(0, 3) == "Req") //Request timed out, eat a couple empty lines and start over
				{
					for (short i = 0; i < 4; i++)
					{
						std::getline(readFromFile, rawInput);
					}
					continue;
				}
				else if (rawInput.substr(0,3) == "Rep") //Reply from IP, parse IP and add to vector
				{
					goodPingVector.push_back(parseIPFromRawString(tempIPLine, std::string(IPBase)));
					for (short i = 0; i < 6; i++)
					{
						std::getline(readFromFile, rawInput);
					}
				}
				else
				{
					return 1;
				}
			}
		}
		readFromFile.close();
	}
	else
	{
		return 2;
	}
	std::ofstream writeToFile;
	writeToFile.open("../tmp/parsedips.txt", std::ofstream::app);
	if (writeToFile.is_open())
	{
		if (goodPingVector.empty())
		{
			writeToFile << '0' << std::endl;
		}
		else
		{
			writeToFile << goodPingVector.size() << std::endl;
			for (std::vector<std::string>::const_iterator iter = goodPingVector.begin(); iter != goodPingVector.end(); iter++)
			{
				writeToFile << *iter << std::endl;
			}
		}
		writeToFile.close();
	}
	else
	{
		return 3;
	}
	return 0;
}

//Private Declare Function cleanUpTmp Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" () As Integer
short __stdcall cleanUpTmp()
{
	short resultsReturn = DeleteFile("../tmp/changeresults.txt");
	short pingResultsReturn = DeleteFile("../tmp/pingresults.txt");
	short pingReturn = DeleteFile("../tmp/tmpping.bat"); 
	short parsedReturn = DeleteFile("../tmp/parsedips.txt");
	return (resultsReturn || pingResultsReturn || pingReturn || parsedReturn);
}

//Private Declare Function delay Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" (ByVal howLong As Double) As Integer
short __stdcall delay(double howLong)
{
	clock_t startTime = clock();
	double elapsedTime = static_cast<double>((clock() - startTime) / (CLOCKS_PER_SEC));
	while (elapsedTime < howLong)
	{
		elapsedTime = static_cast<double>((clock() - startTime) / (CLOCKS_PER_SEC));
	}
	return 0;
}

std::string parseIPFromRawString(const std::string &stringToParse, const std::string &IPBaseArg)
{
	std::string IPBase = IPBaseArg;
	size_t findInString = stringToParse.find(IPBase);
	size_t startOfTail = (findInString + IPBase.length());
	size_t lengthOfTail = 0;
	for (short i = 0; i < 3; i++)
	{
		if (isNumeric(stringToParse[startOfTail + i]))
		{
			lengthOfTail++;
		}
	}
	IPBase.append(stringToParse.substr(startOfTail, lengthOfTail));
	return IPBase;
}
	

bool isNumeric(const char &charToCheck)
{
	return ((charToCheck == '0') || (charToCheck == '1') || (charToCheck == '2') || (charToCheck == '3') || (charToCheck == '4') || (charToCheck == '5') || (charToCheck == '6') || (charToCheck == '7') || (charToCheck == '8') || (charToCheck == '9'));  
}
 

bool directoryExists(const std::string &directoryToCheck)
{
	DWORD ftyp = GetFileAttributesA(directoryToCheck.c_str());
	if (ftyp == INVALID_FILE_ATTRIBUTES)
	{
		return false; //Directory does not exist
	}
	if (ftyp & FILE_ATTRIBUTE_DIRECTORY)
    {
		return true; //Directory exists
	}
	return false; 
}


template<typename T>
bool isNull(T *ptrToCheck)
{
	return (ptrToCheck == NULL);
}

template<typename T>
std::string toString(const T &input)
{
	std::stringstream transfer;
	std::string returnString;
	transfer << input;
	transfer >> returnString;
	return returnString;
}

