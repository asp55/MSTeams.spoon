# MSTeams.spoon
A spoon to interact with the Microsoft Teams local websocket API

# Enable the 3rd party api
As of the August 3, 2026 release (version 26198.202.4929.7171) the Mac version of Teams no longer ships with the Third-Party app API enabled. To enable it create `configuration.json` in `~/Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams`

with a value of 
```
{
  "thirdPartyDevices/thirdPartyDevicesManagerEnabled": true
}
```

and restart teams

# Credits
Microsoft Teams support was able to be added thanks to the [documentation of the local websocket api](https://github.com/svrooij/teams-monitor?tab=readme-ov-file#teams-has-a-local-api) by [@svrooij](https://github.com/svrooij)
