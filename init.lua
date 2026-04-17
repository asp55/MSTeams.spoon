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



--- MSTeams.permissions
--- Variable
--- Read-only table of meetingPermissions provided by Teams or nil if not connected to teams.
local meetingPermissions = nil

--- MSTeams.state
--- Variable
--- Read-only table of meetingState properties provided by Teams or nil if not connected to teams.
local meetingState = nil

local updateCallback = function () end

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

   local requestMsg = hs.json.encode({
      requestId = requestID,
      action = msg.action or "",
      parameters = msg.parameters or {}
   });

   if teamsWebsocket then
      teamsWebsocket:send(requestMsg)
   end
end

local closedCount = 0

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

         updateCallback()
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

--- MSTeams:onUpdate(callback)-> spoon.MSTeams 
--- Method
--- Runs a callback whenever state or permissions are updated
---
--- Parameters:
---  * callback - the callback function to run on update; 
---
--- Returns:
---  * The spoon.MSTeams object
function MSTeams:onUpdate(callback)

   if type(callback)=='function' then
      updateCallback = callback
   else
      error('callback must be a function',3)
   end

   return self
end


-------------------------------------------
-- Teams Command Methods
-------------------------------------------
-- query-state

--- MSTeams:queryState()-> spoon.MSTeams 
--- Method
--- Sends a query-state command to Microsoft Teams. 
--- Teams should then subsequently send a meetingUpdate message
--- with meetingPermissions and, if the spoon has been paired, meetingState 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:queryState()
   if not running then
      MSTeams.logger.w("spoon MSTeams must be running to perform actions. Run MSTeams:start()")
   elseif not teamsWebsocket then
      MSTeams.logger.w("Not connected to teams")
   else
      sendRequest({action="query-state",parameters={}})
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

--- MSTeams:pair()-> spoon.MSTeams 
--- Method
--- If able: sends a pair command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:pair()
   if canAct(permission.canPair) then
      sendRequest({action="pair", parameters={}})
   end

   return self
end


-- toggle-mute

--- MSTeams:toggleMute()-> spoon.MSTeams 
--- Method
--- If able: sends a toggle-mute command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:toggleMute()
   if canAct(permission.canToggleMute) then
      sendRequest({action="toggle-mute", parameters={}})
   end

   return self
end

-- mute

--- MSTeams:mute()-> spoon.MSTeams 
--- Method
--- If able: sends a mute command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:mute()
   if canAct(permission.canToggleMute) then
      sendRequest({action="mute", parameters={}})
   end

   return self
end

-- unmute

--- MSTeams:unmute()-> spoon.MSTeams 
--- Method
--- If able: sends a unmute command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:unmute()
   if canAct(permission.canToggleMute) then
      sendRequest({action="unmute", parameters={}})
   end

   return self
end



-- toggle-video

--- MSTeams:toggleVideo()-> spoon.MSTeams 
--- Method
--- If able: sends a toggle-video command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:toggleVideo()
   if canAct(permission.canToggleVideo) then
      sendRequest({action="toggle-video", parameters={}})
   end

   return self
end

-- show-video

--- MSTeams:showVideo()-> spoon.MSTeams 
--- Method
--- If able: sends a show-video command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:showVideo()
   if canAct(permission.canToggleVideo) then
      sendRequest({action="show-video", parameters={}})
   end

   return self
end

-- hide-video

--- MSTeams:hide-video()-> spoon.MSTeams 
--- Method
--- If able: sends a hide-video command to Microsoft Teams.
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:hideVideo()
   if canAct(permission.canToggleVideo) then
      sendRequest({action="hide-video", parameters={}})
   end

   return self
end


-- stop-sharing

--- MSTeams:stopSharing()-> spoon.MSTeams 
--- Method
--- If able: sends a stop-sharing command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:stopSharing()
   if canAct(permission.canStopSharing) then
      sendRequest({action="stop-sharing", parameters={}})
   end
   return self
end


-- toggle-background-blur

--- MSTeams:toggleBlurBackground()-> spoon.MSTeams 
--- Method
--- If able: sends a toggle-background-blur command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:toggleBlurBackground()
   if canAct(permission.canToggleBlur) then
      sendRequest({action="toggle-background-blur", parameters={}})
   end

   return self
end

-- blur-background

--- MSTeams:blurBackground()-> spoon.MSTeams 
--- Method
--- If able: sends a blur-background command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:blurBackground()
   if canAct(permission.canToggleBlur) then
      sendRequest({action="blur-background", parameters={}})
   end

   return self
end

-- unblur-background

--- MSTeams:unblurBackground()-> spoon.MSTeams 
--- Method
--- If able: sends a unblur-background command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:unblurBackground()
   if canAct(permission.canToggleBlur) then
      sendRequest({action="unblur-background", parameters={}})
   end

   return self
end


-- toggle-hand

--- MSTeams:toggleHand()-> spoon.MSTeams 
--- Method
--- If able: sends a toggle-hand command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:toggleHand()
   if canAct(permission.canToggleHand) then
      sendRequest({action="toggle-hand", parameters={}})
   end

   return self
end

-- raise-hand

--- MSTeams:raiseHand()-> spoon.MSTeams 
--- Method
--- If able: sends a raise-hand command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:raiseHand()
   if canAct(permission.canToggleHand) then
      sendRequest({action="raise-hand", parameters={}})
   end

   return self
end

-- lower-hand

--- MSTeams:lowerHand()-> spoon.MSTeams 
--- Method
--- If able: sends a lower-hand command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:lowerHand()
   if canAct(permission.canToggleHand) then
      sendRequest({action="lower-hand", parameters={}})
   end

   return self
end

-- send-reaction {"type":"like"}

--- MSTeams:reactLike()-> spoon.MSTeams 
--- Method
--- If able: sends a send-reaction command of type:"like" to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:reactLike(reaction)
   if canAct(permission.canReact) then
      sendRequest({action="send-reaction", parameters={type="like"}})
   end

   return self
end


-- send-reaction {"type":"love"}

--- MSTeams:reactLove()-> spoon.MSTeams 
--- Method
--- If able: sends a send-reaction command of type:"love" to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:reactLove(reaction)
   if canAct(permission.canReact) then
      sendRequest({action="send-reaction", parameters={type="love"}})
   end

   return self
end

-- send-reaction {"type":"applause"}

--- MSTeams:reactApplause()-> spoon.MSTeams 
--- Method
--- If able: sends a send-reaction command of type:"applause" to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:reactApplause(reaction)
   if canAct(permission.canReact) then
      sendRequest({action="send-reaction", parameters={type="applause"}})
   end

   return self
end

-- send-reaction {"type":"laugh"}

--- MSTeams:reactLaugh()-> spoon.MSTeams 
--- Method
--- If able: sends a send-reaction command of type:"laugh" to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:reactLaugh(reaction)
   if canAct(permission.canReact) then
      sendRequest({action="send-reaction", parameters={type="laugh"}})
   end

   return self
end

-- send-reaction {"type":"wow"}

--- MSTeams:reactWow()-> spoon.MSTeams 
--- Method
--- If able: sends a send-reaction command of type:"wow" to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:reactWow(reaction)
   if canAct(permission.canReact) then
      sendRequest({action="send-reaction", parameters={type="wow"}})
   end

   return self
end

-- leave-call

--- MSTeams:leaveCall()-> spoon.MSTeams 
--- Method
--- If able: sends a leave-call command to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:leaveCall()
   if canAct(permission.canLeave) then
      sendRequest({action="leave-call", parameters={}})
   end

   return self
end

-- toggle-ui {"type":"chat"}

--- MSTeams:toggleChat()-> spoon.MSTeams 
--- Method
--- If able: sends a toggle-ui command of type:"chat" to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:toggleChat()
   if canAct(permission.canToggleChat) then
         sendRequest({action="toggle-ui", parameters={type="chat"}})
   end
   return self
end

-- toggle-ui {"type":"share-tray"}

--- MSTeams:toggleShareTray()-> spoon.MSTeams 
--- Method
--- If able: sends a toggle-ui command of type:"share-tray" to Microsoft Teams. 
---
--- Parameters:
---  * None
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:toggleShareTray()
   if canAct(permission.canToggleShareTray) then
         sendRequest({action="toggle-ui", parameters={type="share-tray"}})
   end
   return self
end



--- MSTeams:customRequest(action[, parameters])-> spoon.MSTeams 
--- Method
--- If able: sends a custom command to Microsoft Teams. 
---
--- Parameters:
---  * action - string of the action to send to teams
---  * parameters - (optional) table containing parameters to include in the command
---
--- Returns:
---  * The spoon.MSTeams object for chaining
function MSTeams:customRequest(action, parameters)
   if canAct() then
      sendRequest({action=action, parameters=parameters or {}})
   end

   return self
end


-------------------------------------------
-- End of Teams Command Methods
-------------------------------------------

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
