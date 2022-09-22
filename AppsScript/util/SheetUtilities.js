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

function deleteTriggerByScriptName(scriptName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const trigger = triggers[i];
    if (trigger.getHandlerFunction() === scriptName)
      ScriptApp.deleteTrigger(trigger);
  }
}

function getFuncName() {
   return getFuncName.caller.name
}