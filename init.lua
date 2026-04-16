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
local running = false

-------------------------------------------
-- End of Declare Variables
-------------------------------------------



-------------------------------------------
-- Teams Monitor
-------------------------------------------

local teamsPairing = false
local teamsWebsocket = nil
local teamsConnectionId = 0

local function disconnectFromTeams()
   if teamsWebsocket then
      teamsWebsocket:close()
      teamsWebsocket = nil
   end
end

-- forward declare connectToTeams so onTeamsMessage can reference it for reconnects
local connectToTeams = function() end

local requestID = 0
local function sendRequest(msg)
   requestID = requestID + 1
   local requestMsg = string.format('{"requestId":%s, %s}', requestID, msg)
   if teamsWebsocket then
      teamsWebsocket:send(requestMsg)
   end
end

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

      -- if parsed.meetingUpdate and parsed.meetingUpdate.meetingPermissions and parsed.meetingUpdate.meetingPermissions.canPair and not teamsPairing then

      --    MSTeams.logger.d("Sending pairing request")
      --    teamsPairing = true
      --    sendRequest('"action":"pair","parameters":{}')
      -- end

      if parsed.response and parsed.response == "Pairing response resulted in no action" then 
         MSTeams.logger.d("Didn't pair. Will try again next meeting.")
         teamsPairing = false
      end

      if parsed.meetingUpdate and parsed.meetingUpdate.meetingState then
         local ms = parsed.meetingUpdate.meetingState
         
         MSTeams.logger.d("Got new meeting state", hs.inspect.inspect(ms))
      end

   elseif wsType == "closed" then
      teamsWebsocket = nil
      teamsPairing = false
      if running then
         MSTeams.logger.d("Teams WebSocket closed, probably because this app was blocked from the Third-party app API in teams.")
         MSTeams.logger.d("Go to Settings > Privacy > Third-party app API > Manage API and remove the application from block.")
      end

   elseif wsType == "fail" then
      teamsWebsocket = nil
      if running then
         MSTeams.logger.d("Teams not available, retrying in 30 seconds")
         hs.timer.doAfter(30, connectToTeams)
      end
   end
end

connectToTeams = function()
   MSTeams.logger.d("Connect to teams")
   -- Increment the connection ID before closing, so any callbacks from the
   -- previous connection are ignored even if close() fires synchronously.
   teamsConnectionId = teamsConnectionId + 1
   local myId = teamsConnectionId
   if teamsWebsocket then
      teamsWebsocket:close()
      teamsWebsocket = nil
   end
   local token = hs.settings.get("MSTeams.teamsToken") or ""
   local manufacturer = "Hammerspoon"
   local device = "MSTeams.spoon"
   local app = "MSTeams.spoon"
   local url = "ws://localhost:8124?token="..token.."&protocol-version=2.0.0&manufacturer="..manufacturer.."&device="..device.."&app="..app.."&app-version="..MSTeams.version
   MSTeams.logger.d("Connecting to Teams")
   teamsWebsocket = hs.websocket.new(url, function(wsType, message)
      if myId == teamsConnectionId then
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

--- MSTeams:start() -> spoon.MSTeams
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

   if(not running) then
      running = true
      if(not teamsWebsocket) then
         connectToTeams()
      end
   end
 
   return self
end

--- MSTeams:stop()-> spoon.MSTeams 
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
   running = false
   disconnectFromTeams()
   return self
end


MSTeams.actions = {}

-- query-state
function MSTeams.actions:queryState()
   if running then
      sendRequest('"action":"query-state","parameters":{}')
   end

   return self
end

-- pair
function MSTeams.actions:pair()
   if running then
      sendRequest('"action":"pair","parameters":{}')
   end

   return self
end



-- toggle-mute
function MSTeams.actions:toggleMute()
   if running then
      sendRequest('"action":"toggle-mute","parameters":{}')
   end

   return self
end

-- mute
function MSTeams.actions:mute()
   if running then
      sendRequest('"action":"mute","parameters":{}')
   end

   return self
end

-- unmute
function MSTeams.actions:unmute()
   if running then
      sendRequest('"action":"unmute","parameters":{}')
   end

   return self
end



-- toggle-video
function MSTeams.actions:toggleVideo()
   if running then
      sendRequest('"action":"toggle-video","parameters":{}')
   end

   return self
end

-- show-video
function MSTeams.actions:showVideo()
   if running then
      sendRequest('"action":"show-video","parameters":{}')
   end

   return self
end

-- hide-video
function MSTeams.actions:hideVideo()
   if running then
      sendRequest('"action":"hide-video","parameters":{}')
   end

   return self
end


-- stop-sharing
function MSTeams.actions:stopSharing()
   if running then
      sendRequest('"action":"stop-sharing","parameters":{}')
   end
   return self
end


-- toggle-background-blur
function MSTeams.actions:toggleBlurBackground()
   if running then
      sendRequest('"action":"toggle-background-blur","parameters":{}')
   end

   return self
end

-- blur-background
function MSTeams.actions:blurBackground()
   if running then
      sendRequest('"action":"blur-background","parameters":{}')
   end

   return self
end

-- unblur-background
function MSTeams.actions:unblurBackground()
   if running then
      sendRequest('"action":"unblur-background","parameters":{}')
   end

   return self
end


-- toggle-hand
function MSTeams.actions:toggleHand()
   if running then
      sendRequest('"action":"toggle-hand","parameters":{}')
   end

   return self
end

-- raise-hand
function MSTeams.actions:raiseHand()
   if running then
      sendRequest('"action":"raise-hand","parameters":{}')
   end

   return self
end

-- lower-hand
function MSTeams.actions:lowerHand()
   if running then
      sendRequest('"action":"lower-hand","parameters":{}')
   end

   return self
end

local reactionConstants = {like="like", love="love", applause="applause", laugh="laugh" }
-- send-reaction {"type":"like"|"love"|"applause"|"laugh"}

function MSTeams.actions:sendReaction(reaction)
   if running then
      local reactionType = reactionConstants[reaction]
      if reactionType then
         sendRequest('"action":"send-reaction","parameters":{"type":"'..reactionType..'"}')
      else
         --Invalid reaction
      end
   end

   return self
end

-- leave-call
function MSTeams.actions:leaveCall()
   if running then
      sendRequest('"action":"leave-call","parameters":{}')
   end

   return self
end

local uiConstants = {chat="chat", shareTray="share-tray"}
-- toggle-ui {"type":"chat"}

function MSTeams.actions:toggleUI(element)
   if running then
      local uiElement = nil
      for _, v in pairs(uiConstants) do
         if(v==element) then uiElement=element end
      end
      print(uiElement)
      if uiElement then
         sendRequest('"action":"toggle-ui","parameters":{"type":"'..uiElement..'"}')
      else
         --Invalid reaction
      end
   end

   return self
end



function MSTeams.actions:customRequest(msg)
   if running then
      sendRequest(msg)
   end

   return self
end



-------------------------------------------
-- End of Methods
-------------------------------------------

return setmetatable({}, {
   __index=function (_, k)
      if k=="reaction" then
         return reactionConstants
      elseif k=="ui" then
         return uiConstants
      else
         return MSTeams[k]
      end
   end,
   __newindex=function (_, k, v)
      if k=="reaction" or k=="ui" or k=="actions" then
         --skip readonly constants
      else
         MSTeams[k]=v
      end
   end
})
