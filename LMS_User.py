#!/usr/bin/env python
# -*- coding: utf-8 -*-

import configparser
import os
import threading
import time

import pyodbc
import wx


class FramedImage(wx.Panel):
    """带边框的图片类"""
    _primary_image = None  # 原始主图片
    _primary_bitmap = None  # 主图片的位图
    _primary_bitmap_size = None  # 裁剪后Panel的大小
    _frame_image = None  # 原始边框图片
    _frame_bitmap = None  # 边框图片的位图

    def __init__(self, parent, client_size, primary_image_file, frame_image_file=None):
        if isinstance(client_size, wx.Size):
            # 调用Panel基类的初始化函数
            wx.Panel.__init__(self, parent, id=wx.ID_ANY, size=client_size)
        
            if os.path.exists(primary_image_file):
                # 获取原始主图片
                self._primary_image = wx.Image(primary_image_file, wx.BITMAP_TYPE_ANY)
                # 根据框架大小对主图片副本进行缩放调整（图片宽高比不变）
                primary_image = self._primary_image.Copy()
                x1 = client_size[0] * primary_image.GetHeight()
                x2 = client_size[1] * primary_image.GetWidth()
                if x1 < x2:
                    new_height = int(x1 / float(primary_image.GetWidth()))
                    primary_image.Rescale(client_size[0], new_height)
                else:
                    new_width = int(x2 / float(primary_image.GetHeight()))
                    primary_image.Rescale(new_width, client_size[1])
                # 生成主图片的bitmap
                self._primary_bitmap = wx.Bitmap(primary_image)
                # 记录主图片bitmap的大小
                self._primary_bitmap_size = wx.Size(primary_image.GetWidth(), primary_image.GetHeight())
        
            if frame_image_file is not None:
                if os.path.exists(frame_image_file):
                    # 获取原始边框图片，并设置其部分透明
                    self._frame_image = wx.Image(frame_image_file, wx.BITMAP_TYPE_ANY)
                    if not self._frame_image.HasAlpha():
                        self._frame_image.InitAlpha()
                    # 将边框图片副本的大小调整为Panel的大小
                    frame_image = self._frame_image.Copy()
                    frame_image.Rescale(client_size[0], client_size[1])
                    # 生成边框图片的bitmap
                    self._frame_bitmap = wx.Bitmap(frame_image)
            
            self.Show()
            
            self.Bind(wx.EVT_ERASE_BACKGROUND, self._on_erase_background)
            self.Bind(wx.EVT_SIZE, self._on_size)
    
    def _redraw(self, dc):
        dc.Clear()
        # 显示主图片
        if self._primary_bitmap:
            client_size = self.GetClientSize()
            x = int((client_size[0] - self._primary_bitmap_size[0]) / 2.0)
            y = int((client_size[1] - self._primary_bitmap_size[1]) / 2.0)
            dc.DrawBitmap(self._primary_bitmap, x, y)
        # 显示边框图片
        if self._frame_bitmap:
            dc.DrawBitmap(self._frame_bitmap, 0, 0)
    
    def _on_erase_background(self, event):
        dc = wx.ClientDC(self)
        self._redraw(dc)
    
    def _on_size(self, event):
        client_size = self.GetClientSize()
        if self._primary_image is not None:
            # 获取主图片副本，并根据Panel大小对其缩放（保持主图片宽高比不变）
            primary_image = self._primary_image.Copy()
            x1 = client_size[0] * primary_image.GetHeight()
            x2 = client_size[1] * primary_image.GetWidth()
            if x1 < x2:
                new_height = int(x1 / float(primary_image.GetWidth()))
                primary_image.Rescale(client_size[0], new_height)
            else:
                new_width = int(x2 / float(primary_image.GetHeight()))
                primary_image.Rescale(new_width, client_size[1])
            # 生成主图片的bitmap
            self._primary_bitmap = wx.Bitmap(primary_image)
            # 记录主图片位图的大小
            self._primary_bitmap_size = wx.Size(primary_image.GetWidth(), primary_image.GetHeight())

        if self._frame_image is not None:
            # 将边框图片调整为Panel的大小
            frame_image = self._frame_image.Copy()
            frame_image = frame_image.Rescale(client_size[0], client_size[1])
            # 生成边框图片的bitmap
            self._frame_bitmap = wx.Bitmap(frame_image)
        # 发送绘制背景消息
        wx.QueueEvent(self, wx.PyCommandEvent(wx.wxEVT_ERASE_BACKGROUND))
        event.Skip()

    def change_frame(self, frame_image_file):
        if os.path.exists(frame_image_file):
            # 记录新边框图片为当前边框图片，并设置其透明
            self._frame_image = wx.Image(frame_image_file, wx.BITMAP_TYPE_ANY)
            if not self._frame_image.HasAlpha():
                self._frame_image.InitAlpha()
            # 获取当前边框图片的副本
            frame_image = self._frame_image.Copy()
            # 将边框图片副本调整为Panel的大小
            client_size = self.GetClientSize()
            frame_image = frame_image.Rescale(client_size[0], client_size[1])
            # 生成边框图片的bitmap
            self._frame_bitmap = wx.Bitmap(frame_image)
            # 发送绘制背景消息
            wx.QueueEvent(self, wx.PyCommandEvent(wx.wxEVT_ERASE_BACKGROUND))
    
    def change_image(self, primary_image_file):
        if os.path.exists(primary_image_file):
            client_size = self.GetClientSize()
            # 获取主图片，并将其缩放适应Panel大小
            self._primary_image = wx.Image(primary_image_file, wx.BITMAP_TYPE_ANY)
            primary_bitmap = self._primary_image.Copy()
            x1 = client_size[0] * primary_bitmap.GetHeight()
            x2 = client_size[1] * primary_bitmap.GetWidth()
            if x1 < x2:
                new_height = int(x1 / float(primary_bitmap.GetWidth()))
                primary_bitmap.Rescale(client_size[0], new_height)
            else:
                new_width = int(x2 / float(primary_bitmap.GetHeight()))
                primary_bitmap.Rescale(new_width, client_size[1])
            # 生成主图片的位图
            self._primary_bitmap = wx.Bitmap(primary_bitmap)
            # 记录新的主图片位图的大小
            self._primary_bitmap_size = wx.Size(primary_bitmap.GetWidth(), primary_bitmap.GetHeight())
            # 发送绘制背景消息
            wx.QueueEvent(self, wx.PyCommandEvent(wx.wxEVT_ERASE_BACKGROUND))


class UserUI(wx.Frame):
    """用户借还书操作界面"""
    _bg_image = None
    _st_user_id = None
    _cho_user_id = None
    _st_user_class = None
    _cho_user_class = None
    _st_user_name = None
    _user_photo = None
    _book_photo = None
    _st_book_name = None
    _lc_books = None
    _bs_user = None
    _bs_book = None
    _timer_auto_clear_user = None
    _bInit = False
    _input_string = ''

    def __init__(self, parent, title):
        frame_style = wx.CAPTION | wx.CLOSE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU | wx.CLIP_CHILDREN
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=title, pos=wx.DefaultPosition, size=wx.DefaultSize,
                          style=frame_style | wx.TRANSPARENT_WINDOW | wx.WANTS_CHARS)
        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        # 读取配置文件中关于显示的设置
        config = configparser.ConfigParser()
        if os.path.exists(os.getcwd() + '\\LMS.ini'):
            config.read_file(open(os.getcwd() + '\\LMS.ini'))
            user_photo_size = int(config.get("UserUI", "User_Photo_Size"))
            book_photo_size = int(config.get("UserUI", "Book_Photo_Size"))
            book_list_width = int(config.get("UserUI", "Book_List_Width"))
            self._auto_clear_user_time = int(config.get("UserUI", "Auto_Clear_User_Time"))
            self._bInit = True
        else:
            wx.MessageBox("未找到配置文件LMS.ini")
            self.Destroy()
            return
    
        # 背景
        self._bg_image = wx.Image(os.getcwd() + r"\bg.jpg", wx.BITMAP_TYPE_ANY)
        # 用户id和用户班级字体
        font1 = wx.Font(22, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "楷体")
        # 用户姓名和图书名称字体
        font2 = wx.Font(32, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "楷体")
        # 借书列表字体
        font3 = wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "新宋体")
    
        # #########################################界面元素######################################### #
        # # 用户班级
        # self._st_user_class = wx.StaticText(self, wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize,
        #                                     wx.ALIGN_CENTER_HORIZONTAL | wx.TRANSPARENT_WINDOW)
        # self._st_user_class.SetFont(font1)
        # self._st_user_class.Wrap(-1)
        # self._st_user_class.Hide()
        # # 用户ID
        # self._st_user_id = wx.StaticText(self, wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize,
        #                                  wx.ALIGN_CENTER_HORIZONTAL | wx.TRANSPARENT_WINDOW)
        # self._st_user_id.SetFont(font1)
        # self._st_user_id.Wrap(-1)
        # self._st_user_id.Hide()
        # 用户ID
        self._cho_user_id = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, ["用户ID"])
        self._cho_user_id.SetSelection(0)
        self._cho_user_id.SetFont(font1)
        self._cho_user_id.Disable()
        # 用户班级
        self._cho_user_class = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, ["班级"])
        self._get_classes()
        self._cho_user_class.SetSelection(0)
        self._cho_user_class.SetFont(font1)
        self._cho_user_class.Disable()
        # 用户照片
        p_image_file = os.getcwd() + r"\default_user.jpg"
        f_image_file = os.getcwd() + r"\frame_user.png"
        self._user_photo = FramedImage(self, wx.Size(user_photo_size, user_photo_size), p_image_file, f_image_file)
        # 用户姓名
        self._st_user_name = wx.StaticText(self, wx.ID_ANY, u"姓名", wx.DefaultPosition,
                                           wx.Size(user_photo_size, user_photo_size),
                                           wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL)
        self._st_user_name.SetFont(font2)
        self._st_user_name.SetForegroundColour("red")
        # self._st_user_name.SetBackgroundColour("white")
        self._st_user_name.Hide()
        self._st_user_name.Wrap(-1)
    
        # 图书照片
        p_image_file = os.getcwd() + r"\default_book.jpg"
        f_image_file = os.getcwd() + r"\frame_book.png"
        self._book_photo = FramedImage(self, wx.Size(book_photo_size, book_photo_size), p_image_file, f_image_file)
        # 图书名称
        self._st_book_name = wx.StaticText(self, wx.ID_ANY, u"图书", wx.DefaultPosition,
                                           wx.Size(book_photo_size, book_photo_size),
                                           wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL)
        self._st_book_name.SetFont(font2)
        self._st_book_name.SetForegroundColour("red")
        # self._st_book_name.SetBackgroundColour("white")
        self._st_book_name.Hide()
        self._st_book_name.Wrap(-1)
    
        # 借书列表
        list_size = wx.Size(book_list_width, 300)
        self._lc_books = wx.ListCtrl(self, wx.ID_ANY, wx.DefaultPosition, list_size,
                                     wx.LC_HRULES | wx.LC_VRULES | wx.LC_SINGLE_SEL | wx.LC_REPORT)
        self._lc_books.SetFont(font3)
        self._lc_books.SetBackgroundColour("white")
        self._lc_books.InsertColumn(0, '', wx.LIST_FORMAT_CENTER, width=6)
        self._lc_books.InsertColumn(1, '序号', wx.LIST_FORMAT_CENTER, width=80)
        width = int(book_list_width * 0.7)
        self._lc_books.InsertColumn(2, '图书', wx.LIST_FORMAT_CENTER, width=width)
        width = int(book_list_width * 0.3)
        self._lc_books.InsertColumn(3, '位置', wx.LIST_FORMAT_CENTER, width=width)
        # #########################################界面元素######################################### #
    
        # #########################################布局信息######################################### #
        # 用户信息布局
        self._bs_user = wx.BoxSizer(wx.VERTICAL)
        # self._bs_user.Add(self._st_user_class, proportion=0, flag=wx.ALIGN_CENTER_HORIZONTAL | wx.BOTTOM, border=5)
        self._bs_user.Add(self._cho_user_class, proportion=0, flag=wx.ALIGN_CENTER_HORIZONTAL | wx.BOTTOM, border=5)
        self._bs_user.Add(self._user_photo, proportion=1, flag=wx.ALIGN_CENTER_HORIZONTAL, border=5)
        # self._bs_user.Add(self._st_user_id, proportion=0, flag=wx.ALIGN_CENTER_HORIZONTAL | wx.TOP, border=5)
        self._bs_user.Add(self._cho_user_id, proportion=0, flag=wx.ALIGN_CENTER_HORIZONTAL | wx.TOP, border=5)
        # 图书信息布局
        self._bs_book = wx.BoxSizer(wx.VERTICAL)
        self._bs_book.Add(self._book_photo, proportion=1, flag=wx.ALIGN_CENTER_HORIZONTAL, border=5)
        # 整体布局
        gbs = wx.GridBagSizer(0, 0)
        gbs.SetFlexibleDirection(wx.BOTH)
        gbs.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)
        gbs.Add(self._bs_user, wx.GBPosition(0, 0), wx.GBSpan(1, 1),
                wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 20)
        gbs.Add(self._bs_book, wx.GBPosition(0, 2), wx.GBSpan(1, 1),
                wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 20)
        gbs.Add(self._lc_books, wx.GBPosition(1, 0), wx.GBSpan(1, 3),
                wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 20)
        gbs.AddGrowableRow(0)
        gbs.AddGrowableRow(1)
        # 全屏布局
        bs_screen = wx.BoxSizer(wx.VERTICAL)
        bs_screen.AddStretchSpacer()
        bs_screen.Add(gbs, proportion=0, flag=wx.ALIGN_CENTER_HORIZONTAL, border=5)
        bs_screen.AddStretchSpacer()

        self.SetSizerAndFit(bs_screen)
        # #########################################布局信息######################################### #
    
        # 发送绘制背景消息
        # wx.QueueEvent(self, wx.PyCommandEvent(wx.wxEVT_ERASE_BACKGROUND))
        self.Center()
        self.Show()
        self.SetFocus()

        self._bind_event()

    # Virtual event handlers, overide them in your derived class
    def _on_erase_background(self, event):
        self._redraw_background()
        # event.Skip()

    def _on_activate(self, event):
        if self:
            self.SetFocus()
        event.Skip()
    
    def _on_close(self, event):
        if self._timer_auto_clear_user:
            if self._timer_auto_clear_user.isAlive():
                self._timer_auto_clear_user.cancel()
        event.Skip()
        
    def _on_char(self, event):
        key = event.GetUnicodeKey()
        # 只记录数字（用户ID和书籍ISBN只包含数字）
        if key in range(ord('0'), ord('9')+1):
            self._input_string += chr(key)
        # Ctrl+C 关闭
        elif key == wx.WXK_CONTROL_C:
            self._auto_clear_user()
            self.Destroy()
        # Ctrl+F 全屏或退出全屏
        elif key == wx.WXK_CONTROL_F:
            if self.IsFullScreen():
                self.ShowFullScreen(False)
            else:
                self.ShowFullScreen(True)
            self.SetFocus()
        # Ctrl+E 用户班级和Id可输入性切换
        elif key == wx.WXK_CONTROL_E:
            if self._cho_user_class.IsEnabled():
                self._cho_user_class.Disable()
                self._cho_user_id.Disable()
            else:
                self._cho_user_class.Enable(True)
                self._cho_user_id.Enable(True)
                if (self._cho_user_id.GetCount() > 1) and (self._cho_user_id.GetSelection() == 0):
                    if self._cho_user_class.GetSelection() == 0:
                        self._cho_user_id.SetString(0, u"用户ID")
                    else:
                        idx = self._cho_user_class.GetSelection()
                        self._cho_user_id.SetString(0, self._cho_user_class.GetString(idx) + u"用户ID")
        else:
            event.Skip()
        self.SetFocus()

    def _on_key_down(self, event):
        key = event.GetKeyCode()
        # 回车键完成字符串输入
        if key == wx.WXK_RETURN:
            if self._input_string.__len__() == 8:
                self.change_user(self._input_string)
                self._input_string = ''
            elif self._input_string.__len__() == 13:
                self.change_book(self._input_string)
                self._input_string = ''
        # Esc键清空输入字符串
        elif key == wx.WXK_ESCAPE:
            self._input_string = ''
            self.SetFocus()
        else:
            event.Skip()

    def _on_ldb_click(self, event):
        if self.IsFullScreen():
            self.ShowFullScreen(False)
        else:
            self.ShowFullScreen(True)
        self.SetFocus()

    # 自定义函数
    # 绑定事件
    def _bind_event(self):
        self.Bind(wx.EVT_ERASE_BACKGROUND, self._on_erase_background)
        self.Bind(wx.EVT_ACTIVATE, self._on_activate)
        self.Bind(wx.EVT_CLOSE, self._on_close)
        self.Bind(wx.EVT_CHAR, self._on_char)
        self.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
        self.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
        if self._bInit:
            self._cho_user_class.Bind(wx.EVT_CHOICE, self._on_cho_user_class)
            self._cho_user_id.Bind(wx.EVT_CHOICE, self._on_cho_user_id)
            # 双击事件
            # self._st_user_class.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            self._cho_user_class.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            self._user_photo.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            # self._st_user_id.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            self._cho_user_id.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            self._st_user_name.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            self._book_photo.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            self._st_book_name.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            self._lc_books.Bind(wx.EVT_LEFT_DCLICK, self._on_ldb_click)
            # 字符事件
            # self._st_user_class.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            self._cho_user_class.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            self._user_photo.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            # self._st_user_id.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            self._cho_user_id.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            self._st_user_name.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            self._book_photo.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            self._st_book_name.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            self._lc_books.Bind(wx.EVT_CHAR_HOOK, self._on_char)
            # 按键事件
            # self._st_user_class.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
            self._cho_user_class.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
            self._user_photo.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
            # self._st_user_id.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
            self._cho_user_id.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
            self._st_user_name.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
            self._book_photo.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
            self._st_book_name.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
            self._lc_books.Bind(wx.EVT_KEY_DOWN, self._on_key_down)

    # 双缓冲绘制背景
    def _redraw_background(self):
        dc = wx.ClientDC(self)
        client_size = self.GetClientSize()
        bg_image = self._bg_image.Copy()
        bg_image.Rescale(client_size[0], client_size[1])
        bitmap = wx.Bitmap(bg_image)
        dc_mem = wx.MemoryDC()
        dc_mem.SelectObject(bitmap)
        dc.Blit(0, 0, client_size[0], client_size[1], dc_mem, 0, 0)

    def _on_cho_user_id(self, event):
        idx = self._cho_user_id.GetSelection()
        if idx != 0:
            user = self._cho_user_id.GetString(idx)
            user_id = user[user.rfind("(") + 1:user.rfind(")")]
            self.change_user(user_id)
        else:
            self._show_user_image(os.getcwd() + r"\default_user.jpg")
            self._show_book_image(os.getcwd() + r"\default_book.jpg")
            self._st_user_name.SetLabel("")
            self._st_book_name.SetLabel("")
            self._lc_books.DeleteAllItems()
        self.SetFocus()

    def _on_cho_user_class(self, event):
        idx = self._cho_user_class.GetSelection()
        if idx != 0:
            self._get_users()
            self._cho_user_id.SetSelection(0)
            self._show_user_image(os.getcwd() + r"\default_user.jpg")
            self._show_book_image(os.getcwd() + r"\default_book.jpg")
            self._st_user_name.SetLabel("")
            self._st_book_name.SetLabel("")
            self._lc_books.DeleteAllItems()
        else:
            self._auto_clear_user()
        self.SetFocus()

    # 更新用户Id组合框
    def _get_users(self):
        self._cho_user_id.Clear()
        idx = self._cho_user_class.GetSelection()
        if idx == 0:
            self._cho_user_id.Append("用户ID")
            self._cho_user_id.SetSelection(0)
        else:
            user_class = self._cho_user_class.GetString(idx)
            self._cho_user_id.Append(user_class + "用户ID")
            # 在数据库中查找班级用户
            conn = pyodbc.connect("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s\\LMSdb.accdb" % os.getcwd())
            cursor = conn.cursor()
            cursor.execute("SELECT user_name, user_id FROM users WHERE user_class=?", user_class)
            rows = cursor.fetchall()
            for row in rows:
                self._cho_user_id.Append(row[0] + "(" + row[1] + ")")
            cursor.close()
            conn.close()

    # 获取班级组合框内容
    def _get_classes(self):
        self._cho_user_class.Clear()
        self._cho_user_class.Append("班级")
        conn = pyodbc.connect("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s\\LMSdb.accdb" % os.getcwd())
        cursor = conn.cursor()
        cursor.execute("SELECT user_class FROM users GROUP BY user_class")
        rows = cursor.fetchall()
        for row in rows:
            if row[0][-2:] != "毕业":
                self._cho_user_class.Append(row[0])
        cursor.close()
        conn.close()

    # 在用户ID组合框中查找用户ID
    def _find_user(self, user_id):
        for idx in range(1, self._cho_user_id.GetCount()):
            user = self._cho_user_id.GetString(idx)
            if user_id == user[user.rfind("(") + 1:user.rfind(")")]:
                return idx
        return wx.NOT_FOUND

    # 显示用户照片
    def _show_user_image(self, image_file):
        if os.path.exists(image_file):
            # 更新用户照片
            self._user_photo.change_image(image_file)
            # 若当前显示的是用户名字，则修改为显示用户照片
            if self._st_user_name.IsShown():
                self._st_user_name.Hide()
                self._bs_user.Remove(1)
                self._user_photo.Show()
                self._bs_user.Insert(1, self._user_photo, proportion=1, flag=wx.ALIGN_CENTER_HORIZONTAL, border=5)
            # 刷新显示
            self._bs_user.Layout()
            self._bs_book.Layout()

    # 显示用户名字
    def _show_user_name(self, user_name):
        # 更新用户姓名
        self._st_user_name.SetLabel(user_name)
        # 若当前显示的是书籍封面照片，则修改为显示书籍名字
        if self._user_photo.IsShown():
            self._user_photo.Hide()
            self._bs_user.Remove(1)
            self._st_user_name.Show()
            self._bs_user.Insert(1, self._st_user_name, proportion=1, flag=wx.ALIGN_CENTER_HORIZONTAL, border=5)
        # 刷新显示
        self._bs_user.Layout()
        self._bs_book.Layout()

    # 显示图书照片
    def _show_book_image(self, image_file):
        if os.path.exists(image_file):
            # 更新图书照片
            self._book_photo.change_image(image_file)
            # 若当前显示的是图书名字，则修改为显示图书封面照片
            if self._st_book_name.IsShown():
                self._st_book_name.Hide()
                self._bs_book.Remove(0)
                self._book_photo.Show()
                self._bs_book.Insert(0, self._book_photo, proportion=1, flag=wx.ALIGN_CENTER_HORIZONTAL, border=5)
            # 刷新显示
            self._bs_user.Layout()
            self._bs_book.Layout()

    # 显示图书名字
    def _show_book_name(self, book_name):
        # 更新图书名字
        self._st_book_name.SetLabel(book_name)
        # 若当前显示的是图书封面照片，则修改为显示图书名字
        if self._book_photo.IsShown():
            self._book_photo.Hide()
            self._bs_book.Remove(0)
            self._st_book_name.Show()
            self._bs_book.Insert(0, self._st_book_name, proportion=1, flag=wx.ALIGN_CENTER_HORIZONTAL, border=5)
        # 刷新显示
        self._bs_user.Layout()
        self._bs_book.Layout()

    # 借书列表追加内容
    def _append_books_list(self, book_isbn, book_name, book_position):
        item_count = self._lc_books.GetItemCount()
        idx = self._lc_books.InsertItem(item_count, str(item_count))
        # 用于查找
        self._lc_books.SetItem(idx, 0, book_isbn)
        # 序号
        self._lc_books.SetItem(idx, 1, str(item_count + 1))
        # 书名
        self._lc_books.SetItem(idx, 2, book_name)
        # 位置（可能为空）
        if book_position is None:
            self._lc_books.SetItem(idx, 3, '')
        else:
            self._lc_books.SetItem(idx, 3, book_position)

    # 更新借书列表
    def _update_books_list(self, book_isbn, conn=None):
        if book_isbn == '':
            return
        elif conn is None:  # 还书，归还书籍底色变灰
            idx = self._lc_books.FindItem(-1, book_isbn)
            self._lc_books.SetItemBackgroundColour(idx, (200, 200, 200))
            self._lc_books.SetItem(idx, 1, '')
        else:  # 借书，将书籍添加到借书列表中
            idx = self._lc_books.FindItem(-1, book_isbn)
            if idx == -1:  # 借书列表中没有该图书
                # 在数据库books表中查找书籍信息
                cursor_tmp = conn.cursor()
                cursor_tmp.execute("""
                                    SELECT book_isbn, book_name, book_position FROM books
                                    WHERE book_isbn = ?""", book_isbn)
                row_book = cursor_tmp.fetchone()
                # 关闭临时数据库连接
                cursor_tmp.close()
                # 在借书列表中追加书籍
                self._append_books_list(row_book[0], row_book[1], row_book[2])
            else:  # 借书列表中有该图书
                # 恢复列表中该行底色和序号
                self._lc_books.SetItemBackgroundColour(idx, (255, 255, 255))
                self._lc_books.SetItem(idx, 1, str(idx + 1))

    # 变更用户
    def change_user(self, user_id):
        if self._timer_auto_clear_user:
            if self._timer_auto_clear_user.isAlive():
                self._timer_auto_clear_user.cancel()
        # 创建数据库连接
        conn = pyodbc.connect("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s\\LMSdb.accdb" % os.getcwd())
        cursor = conn.cursor()
        cursor_tmp = conn.cursor()
        # 在数据库中查找用户姓名、照片和班级
        cursor.execute("SELECT user_name, user_photo, user_class FROM users WHERE user_id = ?", user_id)
        row = cursor.fetchone()
        if row is None:  # 数据库中未查到该用户
            user_name = '用户\n未登记！'
            user_photo = ''
            user_class = ''
        else:  # 数据库中有该用户
            # 获取用户名
            user_name = row[0]
            # 获取用户班级
            user_class = row[2]
            # 获取用户照片文件名
            if row[1] is None:  # 数据库中未登记用户照片时，分2种情况：1.确实无用户照片；2.默认的用户为照片文件名
                default_photo = os.getcwd() + r'\user_img\\' + user_id + '.jpg'
                if os.path.exists(default_photo):
                    user_photo = default_photo
                else:
                    user_photo = ''
            else:
                user_photo = os.getcwd() + r'\user_img\\' + row[1]
        # 更新用户Class
        # self._st_user_class.SetLabel(user_class)
        if user_class == "":
            self._cho_user_class.SetSelection(0)
        else:
            self._cho_user_class.SetSelection(self._cho_user_class.FindString(user_class))
        self._get_users()
        # 更新用户姓名和照片
        if user_photo == '':  # 用户无照片
            self._show_user_name('\n\n' + user_name)
        else:  # 用户有照片
            self._show_user_image(user_photo)
            self._st_user_name.SetLabel('\n\n' + user_name)
        # 更新用户Id
        # self._st_user_id.SetLabel(user_id)
        idx = self._find_user(user_id)
        if idx != wx.NOT_FOUND:
            self._cho_user_id.SetSelection(idx)
        else:
            self._cho_user_id.SetString(0, "(" + user_id + ")")
            self._cho_user_id.SetSelection(0)

        # 更新书籍封面为空封面
        self._show_book_image(os.getcwd() + r"\default_book.jpg")
        self._st_book_name.SetLabel("")

        # 更新用户借书列表
        self._lc_books.DeleteAllItems()
        # 在数据库中查找用户所借书籍
        cursor.execute("SELECT book_isbn FROM records WHERE user_id = ? ORDER BY id ASC", user_id)
        rows = cursor.fetchall()
        for row in rows:
            # 查找书籍信息
            cursor_tmp.execute("SELECT book_isbn, book_name, book_position FROM books WHERE book_isbn = ?", row[0])
            row_book = cursor_tmp.fetchone()
            # 将该书籍信息添加到列表中
            if row_book is not None:
                self._append_books_list(row_book[0], row_book[1], row_book[2])
        # 关闭数据库指针
        cursor_tmp.close()
        cursor.close()
        # 关闭数据库
        conn.close()
        # 启动定时器（一段时间后自动清除用户信息）
        self._timer_auto_clear_user = threading.Timer(self._auto_clear_user_time, self._auto_clear_user)
        self._timer_auto_clear_user.start()
        self.SetFocus()

    # 变更图书
    def change_book(self, book_isbn):
        # user_id = self._st_user_id.GetLabel()
        # if user_id != '':
        idx = self._cho_user_id.GetSelection()
        if idx != 0:
            user = self._cho_user_id.GetString(idx)
            user_id = user[user.rfind("(") + 1:user.rfind(")")]
            # 关闭定时器
            if self._timer_auto_clear_user:
                if self._timer_auto_clear_user.isAlive():
                    self._timer_auto_clear_user.cancel()
            # 在数据库中查找该书籍
            conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s\\LMSdb.accdb' % os.getcwd())
            cursor = conn.cursor()
            cursor.execute("SELECT book_name, book_photo FROM books WHERE book_isbn = ?", book_isbn)
            row = cursor.fetchone()
            if row is None:  # 未查到该书籍
                self._show_book_name('\n\n\n图书\n未登记')
            else:  # 查到该书籍
                # 显示书籍名称或照片
                if row[1] is None:  # 未记录图书封面照片（2种情况：1.照片文件名为默认的isbn号；2.确实无照片）
                    default_photo = os.getcwd() + r'\book_img\\' + book_isbn + '.jpg'
                    if os.path.exists(default_photo):
                        self._show_book_image(default_photo)
                        self._st_book_name.SetLabel('\n\n' + row[0])
                    else:
                        self._show_book_name('\n\n\n' + row[0])
                else:  # 书籍有封面照片
                    self._show_book_image(os.getcwd() + r'\book_img\\' + row[1])
                    self._st_book_name.SetLabel('\n\n\n' + row[0])
                # 更新用户借书数据
                cursor.execute("SELECT id FROM records WHERE user_id = ? and book_isbn = ?", user_id, book_isbn)
                row = cursor.fetchone()
                if row is None:  # 若未查到记录，则为借书操作
                    # 更新数据库
                    self._db_borrow_book(conn, user_id, book_isbn)
                    # 更新用户借书列表
                    self._update_books_list(book_isbn, conn)
                else:  # 还书操作
                    # 更新数据库
                    self._db_return_book(conn, user_id, book_isbn)
                    # 更新用户借书列表（归还书籍底色变灰）
                    self._update_books_list(book_isbn)
            # 关闭数据库及指针
            cursor.close()
            conn.close()
            # 启动定时器（一段时间后自动清除用户信息）
            self._timer_auto_clear_user = threading.Timer(self._auto_clear_user_time, self._auto_clear_user)
            self._timer_auto_clear_user.start()
        self.SetFocus()

    # 数据库中的借书操作
    @staticmethod
    def _db_borrow_book(conn, user_id, book_isbn):
        cursor_tmp = conn.cursor()
        # 在records表中添加借书记录
        cursor_tmp.execute("""
            INSERT INTO records(user_id, book_isbn)
            VALUES(?, ?)""", user_id, book_isbn)
        conn.commit()
        # 在logs表中添加还书记录
        log_time = time.localtime(time.time())
        log_time = time.strftime("%Y/%m/%d %H:%M:%S", log_time)
        cursor_tmp.execute("""
            INSERT INTO logs(user_id, book_isbn, record_time, is_out)
            VALUES(?, ?, ?, ?)""", user_id, book_isbn, log_time, True)
        conn.commit()
        # 关闭临时数据库指针
        cursor_tmp.close()

    # 数据库中的还书操作
    @staticmethod
    def _db_return_book(conn, user_id, book_isbn):
        cursor_tmp = conn.cursor()
        # 删除records表中的借书记录
        cursor_tmp.execute("DELETE FROM records WHERE user_id = ? AND book_isbn = ?", user_id, book_isbn)
        conn.commit()
        # 在logs表中添加还书记录
        log_time = time.localtime(time.time())
        log_time = time.strftime("%Y/%m/%d %H:%M:%S", log_time)
        cursor_tmp.execute("""
            INSERT INTO logs(user_id, book_isbn, record_time, is_out)
            VALUES(?, ?, ?, ?)""", user_id, book_isbn, log_time, False)
        conn.commit()
        # 关闭临时数据库指针
        cursor_tmp.close()

    # 自动清除界面中的用户及书籍信息
    def _auto_clear_user(self):
        # self._st_user_class.SetLabel('')
        self._cho_user_class.SetSelection(0)
        self._show_user_image(os.getcwd() + r"\default_user.jpg")
        # self._st_user_id.SetLabel('')
        self._cho_user_id.Clear()
        self._cho_user_id.Append("用户ID")
        self._cho_user_id.SetSelection(0)
        self._bs_user.Layout()
        self._show_book_image(os.getcwd() + r"\default_book.jpg")
        self._bs_book.Layout()
        self._st_user_name.SetLabel("")
        self._st_book_name.SetLabel("")
        self._lc_books.DeleteAllItems()
        if self._timer_auto_clear_user:
            if self._timer_auto_clear_user.isAlive():
                self._timer_auto_clear_user.cancel()


if __name__ == '__main__':
    app = wx.App()
    frame = UserUI(None, u"信大四幼图书借阅系统")
    frame.ShowFullScreen(True)
    app.MainLoop()
