Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Windows\Media\chord.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close