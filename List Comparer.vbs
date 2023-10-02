Option Explicit

' Specify the paths to the JSON files
Dim followingFilePath
Dim followersFilePath

followingFilePath = "D:\Others\Codes\following.json"
followersFilePath = "D:\Others\Codes\followers_1.json"

' Create a function to read a file's content
Function ReadFile(filePath)
    Dim objFSO, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(filePath) Then
        Set objFile = objFSO.OpenTextFile(filePath, 1)
        ReadFile = objFile.ReadAll
        objFile.Close
    Else
        ReadFile = ""
    End If
End Function

' Function to extract "value" fields from JSON data and store them in a list
Function ExtractValuesToList(jsonString, list)
    Dim objJSON, objItem

    On Error Resume Next
    Execute "Set objJSON = " & jsonString
    On Error GoTo 0

    If Not IsEmpty(objJSON) Then
        If objJSON.Exists("relationships_following") Then
            Dim followingArray
            Set followingArray = objJSON("relationships_following")

            For Each followingObject In followingArray
                If followingObject.Exists("string_list_data") Then
                    Dim stringListData
                    Set stringListData = followingObject("string_list_data")

                    For Each stringObject In stringListData
                        If stringObject.Exists("value") Then
                            list.Add stringObject("value")
                        End If
                    Next
                End If
            Next
        End If
    End If
End Function

' Create lists for "following" and "followers"
Dim following
Dim followers
Set following = CreateObject("System.Collections.ArrayList")
Set followers = CreateObject("System.Collections.ArrayList")

' Load and extract data from "following.json"
Dim followingData
followingData = ReadFile(followingFilePath)
ExtractValuesToList followingData, following

' Load and extract data from "followers_1.json"
Dim followersData
followersData = ReadFile(followersFilePath)
ExtractValuesToList followersData, followers

' Compare "following" with "followers" and output the difference
Dim notFollowingBack
Set notFollowingBack = CreateObject("System.Collections.ArrayList")

For Each follow In following
    If Not followers.Contains(follow) Then
        notFollowingBack.Add follow
    End If
Next

' Output the users not following back
WScript.Echo "Users not following back:"
For Each user In notFollowingBack
    WScript.Echo user
Next