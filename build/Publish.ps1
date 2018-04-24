$ProgressPreference = 'SilentlyContinue'
7z a EWS.zip .\EWS 
Push-AppveyorArtifact EWS.zip