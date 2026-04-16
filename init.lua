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

local closedCount = 0

local meetingPermissions = nil
local meetingState = nil

local function onTeamsMessage(wsType, message)
   MSTeams.logger.v("Teams WebSocket "..wsType, message)

   if wsType ~= "closed" then
      closedCount = 0
   end

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

      if parsed.meetingUpdate then
         if parsed.meetingUpdate.meetingPermissions then
            meetingPermissions = parsed.meetingUpdate.meetingPermissions
            MSTeams.logger.d("Got new meeting permissions", hs.inspect.inspect(meetingPermissions))
         
            if parsed.meetingUpdate.meetingPermissions.canPair and not teamsPairing then
               MSTeams.logger.d("Sending pairing request")
               teamsPairing = true
               MSTeams:pair()
            end
         end

         if parsed.meetingUpdate.meetingState then
            meetingState = parsed.meetingUpdate.meetingState
            
            MSTeams.logger.d("Got new meeting state", hs.inspect.inspect(meetingState))
         end
      end

      if parsed.response and parsed.response == "Pairing response resulted in no action" then 
         MSTeams.logger.d("Didn't pair. Will try again next meeting.")
         teamsPairing = false
      end

   elseif wsType == "closed" then
      teamsWebsocket = nil
      teamsPairing = false
      if running then
         closedCount = closedCount + 1
         if closedCount > 3 then
            MSTeams.logger.w("Teams WebSocket closed multiple times in a row")
            MSTeams.logger.w("This likely means this app was blocked from the Third-party app API in teams.")
            MSTeams.logger.w("Go to Settings > Privacy > Third-party app API > Manage API and remove the application from block.")
            MSTeams.logger.w("Then restart this spoon.")
            MSTeams:stop()
         else
            MSTeams.logger.i("Teams not available, retrying in 5 seconds")
            hs.timer.doAfter(5, connectToTeams)
         end
      end

   elseif wsType == "fail" then
      teamsWebsocket = nil
      if running then
         MSTeams.logger.i("Teams not available, retrying in 30 seconds")
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
   MSTeams.logger.v("MSTeams:start()")

   if(not running) then
      running = true
      if(not teamsWebsocket) then
         connectToTeams()
      end
   else
      MSTeams.logger.w("MSTeams already started")
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
   MSTeams.logger.v("MSTeams:stop()")

   running = false
   closedCount = 0
   meetingPermissions = nil
   meetingState = nil

   disconnectFromTeams()

   return self
end

--- MSTeams:restart()-> spoon.MSTeams 
--- Method
--- Restarts a MSTeams object
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object
function MSTeams:restart()
   return self:stop():start()
end


-- query-state
function MSTeams:queryState()
   if running then
      sendRequest('"action":"query-state","parameters":{}')
   end

   return self
end

local permission = {
   canLeave = "canLeave",
   canPair = "canPair",
   canReact = "canReact",
   canStopSharing = "canStopSharing",
   canToggleBlur = "canToggleBlur",
   canToggleChat = "canToggleChat",
   canToggleHand = "canToggleHand",
   canToggleMute = "canToggleMute",
   canToggleShareTray = "canToggleShareTray",
   canToggleVideo = "canToggleVideo"
}

local function canAct(requiredPermission)
   if not running then
      MSTeams.logger.w("spoon MSTeams must be running to perform actions. Run MSTeams:start()")
      return false
   elseif not teamsWebsocket then
      MSTeams.logger.w("Not connected to teams")
      return false
   elseif not meetingPermissions then
      MSTeams.logger.w("Not currently in a teams meeting")
      return false
   elseif requiredPermission then
      if not permission[requiredPermission] then
         MSTeams.logger.e("Invalid permission")
         return false
      elseif not meetingPermissions[requiredPermission] then
         MSTeams.logger.w("Don't have necessary permission in current meeting.", requiredPermission.." =", meetingPermissions[requiredPermission])
         return false
      end
   end


   return true
end

-- pair
function MSTeams:pair()
   if canAct(permission.canPair) then
      sendRequest('"action":"pair","parameters":{}')
   end

   return self
end


-- toggle-mute
function MSTeams:toggleMute()
   if canAct(permission.canToggleMute) then
      sendRequest('"action":"toggle-mute","parameters":{}')
   end

   return self
end

-- mute
function MSTeams:mute()
   if canAct(permission.canToggleMute) then
      sendRequest('"action":"mute","parameters":{}')
   end

   return self
end

-- unmute
function MSTeams:unmute()
   if canAct(permission.canToggleMute) then
      sendRequest('"action":"unmute","parameters":{}')
   end

   return self
end



-- toggle-video
function MSTeams:toggleVideo()
   if canAct(permission.canToggleVideo) then
      sendRequest('"action":"toggle-video","parameters":{}')
   end

   return self
end

-- show-video
function MSTeams:showVideo()
   if canAct(permission.canToggleVideo) then
      sendRequest('"action":"show-video","parameters":{}')
   end

   return self
end

-- hide-video
function MSTeams:hideVideo()
   if canAct(permission.canToggleVideo) then
      sendRequest('"action":"hide-video","parameters":{}')
   end

   return self
end


-- stop-sharing
function MSTeams:stopSharing()
   if canAct(permission.canStopSharing) then
      sendRequest('"action":"stop-sharing","parameters":{}')
   end
   return self
end


-- toggle-background-blur
function MSTeams:toggleBlurBackground()
   if canAct(permission.canToggleBlur) then
      sendRequest('"action":"toggle-background-blur","parameters":{}')
   end

   return self
end

-- blur-background
function MSTeams:blurBackground()
   if canAct(permission.canToggleBlur) then
      sendRequest('"action":"blur-background","parameters":{}')
   end

   return self
end

-- unblur-background
function MSTeams:unblurBackground()
   if canAct(permission.canToggleBlur) then
      sendRequest('"action":"unblur-background","parameters":{}')
   end

   return self
end


-- toggle-hand
function MSTeams:toggleHand()
   if canAct(permission.canToggleHand) then
      sendRequest('"action":"toggle-hand","parameters":{}')
   end

   return self
end

-- raise-hand
function MSTeams:raiseHand()
   if canAct(permission.canToggleHand) then
      sendRequest('"action":"raise-hand","parameters":{}')
   end

   return self
end

-- lower-hand
function MSTeams:lowerHand()
   if canAct(permission.canToggleHand) then
      sendRequest('"action":"lower-hand","parameters":{}')
   end

   return self
end

-- send-reaction {"type":"like"}
function MSTeams:reactLike(reaction)
   if canAct(permission.canReact) then
      sendRequest('"action":"send-reaction","parameters":{"type":"like"}')
   end

   return self
end


-- send-reaction {"type":"love"}
function MSTeams:reactLove(reaction)
   if canAct(permission.canReact) then
      sendRequest('"action":"send-reaction","parameters":{"type":"love"}')
   end

   return self
end

-- send-reaction {"type":"applause"}
function MSTeams:reactApplause(reaction)
   if canAct(permission.canReact) then
      sendRequest('"action":"send-reaction","parameters":{"type":"applause"}')
   end

   return self
end

-- send-reaction {"type":"laugh"}
function MSTeams:reactLaugh(reaction)
   if canAct(permission.canReact) then
      sendRequest('"action":"send-reaction","parameters":{"type":"laugh"}')
   end

   return self
end

-- leave-call
function MSTeams:leaveCall()
   if canAct(permission.canLeave) then
      sendRequest('"action":"leave-call","parameters":{}')
   end

   return self
end

-- toggle-ui {"type":"chat"}
function MSTeams:toggleChat()
   if canAct(permission.canToggleChat) then
         sendRequest('"action":"toggle-ui","parameters":{"type":"chat"}')
   end
   return self
end

-- toggle-ui {"type":"share tray"}
function MSTeams:toggleShareTray()
   if canAct(permission.canToggleShareTray) then
         sendRequest('"action":"toggle-ui","parameters":{"type":"share tray"}')
   end
   return self
end


function MSTeams:customRequest(msg)
   if canAct() then
      sendRequest(msg)
   end

   return self
end



-------------------------------------------
-- End of Methods
-------------------------------------------

return setmetatable({}, {
   __index=function (_, k)
      if k=="state" then
         return meetingState
      elseif k=="permissions" then
         return meetingPermissions
      else
         return MSTeams[k]
      end
   end,
   __newindex=function (_, k, v)
      if k=="state" or k=="permissions" then
         --skip readonly fields
      else
         MSTeams[k]=v
      end
   end
})
