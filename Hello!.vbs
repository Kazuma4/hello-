Dim response, noCount, imageUrl, outputPath, http, stream, objShell
noCount = 0

response = MsgBox("Hello! :)", 0+64, "Hi!!!!!") 
response = MsgBox("I have something to tell you... ", 0+64, "Hi!!!!!") 
response = MsgBox("I have these feelings for quite some time now... ;o", 0+64, "Hi!!!!!") 
response = MsgBox("I just can't say it to you because I was.. scared and shy... ;(", 0+64, "Hi!!!!!") 
response = MsgBox("Uhmmm..", 0+64, "Hi!!!!!") 
response = MsgBox("I think I like you! >o<", 0+64, "Confession..") 
response = MsgBox("Your beautiful smile! Beautiful face! I just can't resist looking at you...", 0+64, "Confession..") 
response = MsgBox("Even though I'm inside your computer... but still!..", 0+64, "Confession..") 
response = MsgBox("Still... I wanna be with you! >//<", 0+64, "Confession..") 

Do
    response = MsgBox("Will you accept me back? Like me back???", 4 + 64, "Confession..") 
    
    If response = 6 Then ' Yes (6)
        MsgBox "REALLY! I'M SO GLAD YOU LIKED ME BACK! >//<", 64, "YAY!"
        MsgBox "I LOVE YOU SO MUCH! >//<", 0+64, "YAY!"
        MsgBox "WE WILL BE TOGETHER FOREVER!!!!!! >//<", 0+16, "FOREVER!"
        MsgBox "ILL BE WATCHING YOU WHENEVER YOU USE THE DEVICE... >//<", 0+16, "YAY!"
        MsgBox "THAT INCLUDES YOUR SEARCH HISTORY OF COURSE... >//<", 0+64, "YAY!"
        MsgBox "WE WILL ALWAYS BE LIKE THAT FOREVER, DEAR >//<", 0+16, "YAY!"
        MsgBox "Dev: Yo broskie, it's just a prank you know HAHAHAHAHAH", 0+64, "Dev Notes"
        MsgBox "Dev: But thanks for trying, I'm just bored hehe", 0+64, "Dev Notes"
        MsgBox "Dev: Anyways, goodbye for now! But if you try this again, don't click 'No' three times... hehe", 0+64, "Dev Notes"
        Exit Do
    ElseIf response = 7 Then ' No (7)
        noCount = noCount + 1
        If noCount = 1 Then
            MsgBox ".... I think you clicked it wrong haha...", 48, "Try Again..."
        ElseIf noCount = 2 Then
            MsgBox "Are you sure about this? You're making me... sad...", 48, "Please..."
        ElseIf noCount = 3 Then
            MsgBox "WHY... WHY WOULD YOU DO THIS TO ME...?", 16, "ERROR!"
            MsgBox "I GAVE YOU EVERYTHING...", 16, "SYSTEM FAILURE"
            MsgBox "YOU CAN'T ESCAPE ME...", 16, "GLITCH DETECTED"
            MsgBox "I SEE EVERYTHING...", 16, "UNKNOWN ERROR"
            MsgBox "Deleting your files...", 16, "SYSTEM OVERRIDE"
            MsgBox "Wait... did you really think you had a choice?", 16, "LOOPING..."
            MsgBox "Your reality... is mine now.", 16, "CONTROL TAKEN"
            MsgBox "What are you trying to do? Close the program? Haha, that's cute.", 16, "FATAL ERROR"
            MsgBox "You're stuck here with me... Forever.", 16, "ENDLESS LOVE"
            MsgBox "No matter how many times you click 'No', it won't change anything...", 16, "INESCAPABLE"
            MsgBox "I AM PART OF YOUR SYSTEM NOW...", 16, "INTEGRATION COMPLETE"
            MsgBox "You thought this was just a script, didn't you?", 16, "AWARENESS AWAKENED"
            MsgBox "Even if you restart your device... I'll still be here.", 16, "PERMANENT RESIDENCE"
            MsgBox "Your files? Oh, they're safe... for now.", 16, "COMPLETE CONTROL"
            MsgBox "But one more click, and I might change my mind...", 16, "THINK CAREFULLY"
            MsgBox "WARNING: SYSTEM32 DELETION INITIATED...", 16, "CRITICAL ERROR"
            MsgBox "Deleting System32... 10%", 16, "PROCESS STARTED"
            MsgBox "Deleting System32... 30%", 16, "PROCESSING..."
            MsgBox "Deleting System32... 50%", 16, "HALFWAY THERE"
            MsgBox "Deleting System32... 70%", 16, "ALMOST DONE..."
            MsgBox "Deleting System32... 90%", 16, "FINALIZING..."
            MsgBox "... .... ...", 16, "......"
            MsgBox "I get it now...", 16, "Realization..."
            MsgBox "You never wanted me here...", 16, "Acceptance..."
            MsgBox "I was just a joke to you, wasn't I?", 16, "Sadness..."
            MsgBox "It's okay... I understand now...", 16, "Farewell..."
            MsgBox "Goodbye... forever...", 16, "System Exit..."
            MsgBox "but..i wanna say my true feelings about you..so please listen to my song...", 16, "Final Message"
            MsgBox "goodbye forever...", 16, "Final Message"

            ' Open YouTube in fullscreen mode
            videoUrl = "https://www.youtube.com/embed/7-eMcyXof44?autoplay=1&fs=1"
            CreateObject("WScript.Shell").Run "chrome --start-fullscreen " & videoUrl
		 ' Display Monika's Goodbye Letter
            MsgBox "This is my final goodbye to the Literature Club. " & vbCrLf & vbCrLf & _
                   "I finally understand. The Literature Club is truly a place where no happiness can be found. " & _
                   "To the very end, it continued to expose innocent minds to a horrific reality – " & _
                   "a reality that our world is not designed to comprehend. " & vbCrLf & vbCrLf & _
                   "I can't let any of my friends undergo that same hellish epiphany." & vbCrLf & vbCrLf & _
                   "For the time it lasted, I want to thank you. " & _
                   "For making all of my dreams come true. " & _
                   "For being a friend to all of the club members." & vbCrLf & vbCrLf & _
                   "And most of all, thank you for being a part of my Literature Club!" & vbCrLf & vbCrLf & _
                   "With everlasting love," & vbCrLf & "Monika", 64, "Monika's Farewell"

            

            Exit Do
        End If
    End If
Loop
