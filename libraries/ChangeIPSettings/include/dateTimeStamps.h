#ifndef DATETIMESTAMPS_H
#define DATETIMESTAMPS_H

#include <string>

namespace calender
{
    std::string getCurrentYear();
    std::string getCurrentMonth();
	std::string getCurrentMonthName();
    std::string getCurrentDay();
    std::string getCurrentHour24();
    std::string getCurrentHour12();
    std::string getCurrentMinute();
    std::string getCurrentSecond();
    std::string getDateStampMDY();
    std::string getDateStampDMY();
    std::string getTimeStamp12();
    std::string getTimeStamp24();
    std::string getOverallStampMDY12();
    std::string getOverallStampDMY12();
    std::string getOverallStampMDY24();
    std::string getOverallStampDMY24();
}

#endif //dateTimeStamps.h
