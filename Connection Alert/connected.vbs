Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Windows\Media\tada.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close