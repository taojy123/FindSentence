#Boa:Frame:Frame1

import wx
import re
import zipfile
import random

def create(parent):
    return Frame1(parent)

[wxID_FRAME1, wxID_FRAME1BUTTON1, wxID_FRAME1BUTTON2, wxID_FRAME1PANEL1, 
 wxID_FRAME1STATICTEXT1, wxID_FRAME1STATICTEXT2, wxID_FRAME1TEXTCTRL1, 
] = [wx.NewId() for _init_ctrls in range(7)]

class Frame1(wx.Frame):
    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt,
              pos=wx.Point(473, 272), size=wx.Size(342, 271),
              style=wx.DEFAULT_FRAME_STYLE, title='Frame1')
        self.SetClientSize(wx.Size(326, 233))

        self.panel1 = wx.Panel(id=wxID_FRAME1PANEL1, name='panel1', parent=self,
              pos=wx.Point(0, 0), size=wx.Size(326, 233),
              style=wx.TAB_TRAVERSAL)

        self.staticText1 = wx.StaticText(id=wxID_FRAME1STATICTEXT1,
              label=u'Keyword:', name='staticText1', parent=self.panel1,
              pos=wx.Point(48, 80), size=wx.Size(47, 13), style=0)

        self.textCtrl1 = wx.TextCtrl(id=wxID_FRAME1TEXTCTRL1, name='textCtrl1',
              parent=self.panel1, pos=wx.Point(128, 76), size=wx.Size(168, 21),
              style=0, value=u'')

        self.button1 = wx.Button(id=wxID_FRAME1BUTTON1, label=u'Open',
              name='button1', parent=self.panel1, pos=wx.Point(64, 144),
              size=wx.Size(75, 23), style=0)
        self.button1.Bind(wx.EVT_BUTTON, self.OnButton1Button,
              id=wxID_FRAME1BUTTON1)

        self.button2 = wx.Button(id=wxID_FRAME1BUTTON2, label=u'Search',
              name='button2', parent=self.panel1, pos=wx.Point(184, 144),
              size=wx.Size(75, 23), style=0)
        self.button2.Bind(wx.EVT_BUTTON, self.OnButton2Button,
              id=wxID_FRAME1BUTTON2)

        self.staticText2 = wx.StaticText(id=wxID_FRAME1STATICTEXT2,
              label=u'Pleace chose file to open', name='staticText2',
              parent=self.panel1, pos=wx.Point(48, 24), size=wx.Size(240, 13),
              style=0)

    def __init__(self, parent):
        self._init_ctrls(parent)

    def OnButton1Button(self, event):
        self.staticText2.SetLabel(wx.FileSelector("Choose file"))
        event.Skip()

    def OnButton2Button(self, event):
                
        keyword = self.textCtrl1.GetValue().encode("gbk")
        filename = self.staticText2.GetLabel()
        
        a = ""
        if filename[-4:] == "docx":
            z = zipfile.ZipFile(filename)
            a = z.read("word/document.xml")
            z.close()
            a = re.sub(r"<.*?>", '', a) 
            a = a.decode("utf8").encode("gbk", "ignore")
        elif filename[-4:] == ".txt":
            a = open(filename).read()
        
        if a and keyword:
            s = ""
            sentences = a.strip().split('\xa1\xf9') # "\xe2\x80\xbb" for utf8
            for sentence in sentences:
                sentence = sentence.strip()
                if keyword.lower() in sentence.lower():
                    # print sentence
                    s += sentence + "\n"

            open("%s.txt" % keyword, "w").write(s)
            
            wx.MessageBox("The result output to %s.txt" % keyword)
            
        event.Skip()
