/************************************************************************************************************
Krahmer Account Register
Copyright 2022 Douglas Krahmer

This file is part of Krahmer Account Register.

Krahmer Account Register is free software: you can redistribute it and/or modify it under the terms of the 
GNU General Public License as published by the Free Software Foundation, either version 3 of the License, 
or (at your option) any later version.

Krahmer Account Register is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with Krahmer Account Register.
If not, see <https://www.gnu.org/licenses/>.
************************************************************************************************************/

function getNextDateUtc(dateUtc, recurrenceType) {
  switch (recurrenceType)
  {
    case "Biannually":
      return new Date(Date.UTC(dateUtc.getUTCFullYear() + 2, dateUtc.getUTCMonth(), dateUtc.getUTCDate()));
      
    case "Annually":
      return new Date(Date.UTC(dateUtc.getUTCFullYear() + 1, dateUtc.getUTCMonth(), dateUtc.getUTCDate()));
      
    case "Semiannually":
      return new Date(Date.UTC(dateUtc.getUTCFullYear(), dateUtc.getUTCMonth() + 6, dateUtc.getUTCDate()));
      
    case "Quarterly":
      return new Date(Date.UTC(dateUtc.getUTCFullYear(), dateUtc.getUTCMonth() + 3, dateUtc.getUTCDate()));
     
    case "Bimonthly":
      return new Date(Date.UTC(dateUtc.getUTCFullYear(), dateUtc.getUTCMonth() + 2, dateUtc.getUTCDate()));
      
    case "Monthly":
      return new Date(Date.UTC(dateUtc.getUTCFullYear(), dateUtc.getUTCMonth() + 1, dateUtc.getUTCDate()));
      
    case "Semimonthly":
      // Alternate 1st & 15th of the month or any other days 14 days apart. Examples: 12th & 26th, 14th & 28th
      let day = dateUtc.getUTCDate();
      if (day >= 15)
      {
        if (day > 28)
          day = 28; // Must account for the shortest month possible (Feb)
          
          return new Date(Date.UTC(dateUtc.getUTCFullYear(), dateUtc.getUTCMonth() + 1, day - 14));
      }
      else
      {
        return new Date(Date.UTC(dateUtc.getUTCFullYear(), dateUtc.getUTCMonth(), day + 14));
      }
      
    case "Biweekly":
      return new Date(Date.UTC(dateUtc.getUTCFullYear(), dateUtc.getUTCMonth(), dateUtc.getUTCDate() + 14));
      
    case "Weekly":
      return new Date(Date.UTC(dateUtc.getUTCFullYear(), dateUtc.getUTCMonth(), dateUtc.getUTCDate() + 7));
      
    default:
      return null;
  }
}
