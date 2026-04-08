--- === MSTeams ===
---
--- A Spoon to interact with Microsoft Teams via the local websocket API


--We'll store some stuff in an internal table

local _internal = {}


-- create a namespace

local MSTeams={}
MSTeams.__index = MSTeams


-- Metadata
MSTeams.name = "MSTeams"
MSTeams.version = "1.0.0"
MSTeams.author = "Andrew Parnell <aparnell@gmail.com>"
MSTeams.homepage = "https://github.com/asp55/MSTeams.spoon"
MSTeams.license = "MIT - https://opensource.org/licenses/MIT"



-------------------------------------------
-- Declare Variables
-------------------------------------------


--- MSTeams.logger
--- Variable
--- hs.logger object used within the Spoon. Can be accessed to set the default log level for the messages coming from the Spoon.
MSTeams.logger = hs.logger.new('MSTeams')


-- private variable to track if spoon is already running or not. (Makes it easier to find local variables)
_internal.running = false

-------------------------------------------
-- End of Declare Variables
-------------------------------------------



-------------------------------------------
-- Teams Monitor
-------------------------------------------

_internal.teamsInMeeting = false
_internal.teamsWebsocket = nil
_internal.teamsConnectionId = 0

local function disconnectFromTeams()
   if _internal.teamsWebsocket then
      _internal.teamsWebsocket:close()
      _internal.teamsWebsocket = nil
   end
end

-- forward declare connectToTeams so onTeamsMessage can reference it for reconnects
local connectToTeams = function() end

local function onTeamsMessage(wsType, message)
   MSTeams.logger.d("Teams WebSocket "..wsType, message)

   if wsType == "open" then
      MSTeams.logger.d("Connected to Teams local API")

   elseif wsType == "received" then
      local ok, parsed = pcall(hs.json.decode, message)
      if not ok then
         MSTeams.logger.w("Failed to parse Teams message: "..message)
         return
      end

      if parsed.tokenRefresh then
         MSTeams.logger.d("Teams token refreshed")
         hs.settings.set("MSTeams.teamsToken", parsed.tokenRefresh)
      end

      if parsed.meetingUpdate and parsed.meetingUpdate.meetingPermissions and parsed.meetingUpdate.meetingPermissions.canPair and not _internal.teamsPairing then

         MSTeams.logger.d("Sending pairing request")
         _internal.teamsPairing = true
         _internal.teamsWebsocket:send('{"action":"toggle-mute","parameters":{},"requestId":1}')
      end

      if parsed.response and parsed.response == "Pairing response resulted in no action" then 
         MSTeams.logger.d("Didn't pair. Will try again next meeting.")
         _internal.teamsPairing = false
      end

      if parsed.meetingUpdate and parsed.meetingUpdate.meetingState then
         local ms = parsed.meetingUpdate.meetingState
         
         MSTeams.logger.d("Got new meeting state", hs.inspect.inspect(ms))
      end

   elseif wsType == "closed" then
      _internal.teamsWebsocket = nil
      _internal.teamsPairing = false
      if _internal.running then
         MSTeams.logger.d("Teams WebSocket closed, probably because this app was blocked from the Third-party app API in teams.")
         MSTeams.logger.d("Go to Settings > Privacy > Third-party app API > Manage API and remove the application from block.")
      end

   elseif wsType == "fail" then
      _internal.teamsWebsocket = nil
      if _internal.running then
         MSTeams.logger.d("Teams not available, retrying in 30 seconds")
         hs.timer.doAfter(30, connectToTeams)
      end
   end
end

connectToTeams = function()
   MSTeams.logger.d("Connect to teams")
   -- Increment the connection ID before closing, so any callbacks from the
   -- previous connection are ignored even if close() fires synchronously.
   _internal.teamsConnectionId = _internal.teamsConnectionId + 1
   local myId = _internal.teamsConnectionId
   if _internal.teamsWebsocket then
      _internal.teamsWebsocket:close()
      _internal.teamsWebsocket = nil
   end
   local token = hs.settings.get("MSTeams.teamsToken") or ""
   local manufacturer = "Hammerspoon"
   local device = "MSTeams.spoon"
   local app = "MSTeams.spoon"
   local url = "ws://localhost:8124?token="..token.."&protocol-version=2.0.0&manufacturer="..manufacturer.."&device="..device.."&app="..app.."&app-version="..MSTeams.version
   MSTeams.logger.d("Connecting to Teams")
   _internal.teamsWebsocket = hs.websocket.new(url, function(wsType, message)
      if myId == _internal.teamsConnectionId then
         onTeamsMessage(wsType, message)
      end
   end)
end

-------------------------------------------
-- End of Teams Monitor
-------------------------------------------



-------------------------------------------
-- Methods
-------------------------------------------

--- MSTeams:start() -> MSTeams
--- Method
--- Starts a MSTeams object
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object
function MSTeams:start()
   MSTeams.logger.d("Start")

   if(not _internal.running) then
      _internal.running = true
      if(not _internal.teamsWebsocket) then
         connectToTeams()
      end
   end
 
   return self
end

--- MSTeams:stop()
--- Method
--- Stops a MSTeams object
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object
function MSTeams:stop()
   MSTeams.logger.d("Stop")
   _internal.running = false
   disconnectFromTeams()
   return self
end


-------------------------------------------
-- End of Methods
-------------------------------------------

return MSTeams
