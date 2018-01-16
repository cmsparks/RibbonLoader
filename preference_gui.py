import wx

app = wx.App(False)
frame = wx.Frame(None, wx.ID_ANY, "Hello Word!")

frame.Show(True)
app.MainLoop()
