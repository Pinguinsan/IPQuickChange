#include <iostream>
#include <ctime>
#include <locale>
#include <sstream>
#include "../include/dateTimeStamps.h"

namespace calender
{

    std::string getCurrentYear()
    {
        time_t t = time(0);
        struct tm *now = localtime(&t);
        int currentYear = (now->tm_year) + 1900;
        std::stringstream convert;
        std::string returnYear;
        convert << currentYear;
        convert >> returnYear;
        return returnYear;
    }

    std::string getCurrentMonth()
    {
        time_t t = time(0);
        struct tm *now = localtime(&t);
        int currentMonth = (now->tm_mon) + 1;
        std::stringstream convert;
        std::string returnMonth;
        convert << currentMonth;
        convert >> returnMonth;
        return returnMonth;
    }

	std::string getCurrentMonthName()
	{
		time_t t = time(0);
        struct tm *now = localtime(&t);
        int currentMonth = (now->tm_mon) + 1;
		std::string returnMonth = "";
		switch (currentMonth)
		{
			case 1:
			{
				returnMonth.append("January");
				break;
			}
			case 2:
			{
				returnMonth.append("February");
				break;
			}
			case 3:
			{
				returnMonth.append("March");
				break;
			}
			case 4:
			{
				returnMonth.append("April");
				break;
			}			
			case 5:
			{
				returnMonth.append("May");
				break;
			}
			case 6:
			{
				returnMonth.append("June");
				break;
			}
			case 7:
			{
				returnMonth.append("July");
				break;
			}
			case 8:
			{
				returnMonth.append("August");
				break;
			}
			case 9:
			{
				returnMonth.append("September");
				break;
			}
			case 10:
			{
				returnMonth.append("October");
				break;
			}
			case 11:
			{
				returnMonth.append("November");
				break;
			}
			case 12:
			{
				returnMonth.append("December");
				break;
			}
		}
		return returnMonth;
	}

    std::string getCurrentDay()
    {
        time_t t = time(0);
        struct tm *now = localtime(&t);
        int currentDay = now->tm_mday;
        std::stringstream convert;
        std::string returnDay;
        convert << currentDay;
        convert >> returnDay;
        if (currentDay < 10)
        {
            returnDay.insert(returnDay.begin(), '0');
        }
        return returnDay;
    }

    std::string getCurrentHour12()
    {
        std::string AMPM;
        time_t t = time(0);
        struct tm *now = localtime(&t);
        int currentHour = now->tm_hour;
        if (currentHour > 24)
        {
            AMPM = "PM";
            currentHour -= 12;
        }
        else
        {
            AMPM = "AM";
        }
        std::stringstream convert;
        std::string returnHour;
        convert << currentHour;
        convert >> returnHour;
        if (currentHour < 10)
        {
            returnHour.insert(returnHour.begin(), '0');
        }
        return returnHour;
    }

    std::string getCurrentHour24()
    {
        time_t t = time(0);
        struct tm *now = localtime(&t);
        int currentHour = now->tm_hour;
        std::stringstream convert;
        std::string returnHour;
        convert << currentHour;
        convert >> returnHour;
        if (currentHour < 10)
        {
            returnHour.insert(returnHour.begin(), '0');
        }
        return returnHour;
    }

    std::string getCurrentMinute()
    {
        time_t t = time(0);
        struct tm *now = localtime(&t);
        int currentMinute = now->tm_min;
        std::stringstream convert;
        std::string returnMinute;
        convert << currentMinute;
        convert >> returnMinute;
        if (currentMinute < 10)
        {
            returnMinute.insert(returnMinute.begin(), '0');
        }
        return returnMinute;
    }

    std::string getCurrentSecond()
    {
        time_t t = time(0);
        struct tm *now = localtime(&t);
        int currentSecond = now->tm_sec;
        std::stringstream convert;
        std::string returnSecond;
        convert << currentSecond;
        convert >> returnSecond;
        if (currentSecond < 10)
        {
            returnSecond.insert(returnSecond.begin(), '0');
        }
        return returnSecond;
    }

    std::string getTimeStamp12()
    {
        time_t t = time(0);
        struct tm *now = localtime(&t);
        int tempHour = now->tm_hour;
        if (tempHour > 12)
        {
            tempHour += 100;
        }

        std::string returnHour;
        std::string AMPM;
        if (tempHour >= 100)
        {
            AMPM = "PM";
            tempHour -= 112;
        }
        else
        {
            AMPM = "AM";
        }
        std::stringstream convert;
        convert << tempHour;
        convert >> returnHour;
        if (tempHour < 10)
        {
            returnHour.insert(returnHour.begin(), '0');
        }
        return (returnHour + ':' + getCurrentMinute() + ':' + getCurrentSecond() + AMPM);
    }

    std::string getTimeStamp24()
    {
        return (getCurrentHour24() + ':' + getCurrentMinute() + ':' + getCurrentDay());
    }

    std::string getDateStampMDY()
    {
        return (getCurrentMonth() + '-' + getCurrentDay() + '-' + getCurrentYear());
    }

    std::string getDateStampDMY()
    {
        return (getCurrentDay() + '-' + getCurrentMonth() + '-' + getCurrentYear());
    }

    std::string getOverallStampDMY12()
    {
        return (getDateStampDMY() + ' ' + getTimeStamp12());
    }

    std::string getOverallStampDMY24()
    {
        return (getDateStampDMY() + ' ' + getTimeStamp24());
    }

    std::string getOverallStampMDY12()
    {
        return (getDateStampMDY() + ' ' + getTimeStamp12());
    }

    std::string getOverallStampMDY24()
    {
        return (getDateStampMDY() + ' ' + getTimeStamp24());
    }
}
