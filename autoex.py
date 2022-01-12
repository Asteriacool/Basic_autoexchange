from pywinauto import application
from pywinauto import mouse
import time
import psutil


def proc_exist(process_name):
    p1 = psutil.pids()
    for pid in p1:
        if psutil.Process(pid).name() == process_name:
            return pid


def get_app():
    if isinstance(proc_exist('Exchanger.exe'), int):
        # 如果正在运行
        app = application.Application(backend='uia').connect(path='C:/Program Files/CAD Exchanger/bin/exchanger')
    else:
        # 如果没有运行
        app = application.Application(backend='uia').start('C:/Program Files/CAD Exchanger/bin/exchanger')
    return app


# 判断文件列表中是否存在某个文件
# title = None,  # 控件的标题文字，对应inspect中Name字段
def document_is_exist(window, name):
    element = window.descendants(control_type="ListItem", title=name)
    return len(element) != 0


# 判断是否有重命名弹窗
# 如果有弹窗，就把原有的文件覆盖
def controlwindow_is_exist(window):
    element = window.child_window(title="确认另存为", control_type="Window")
    if element.exists():
        win_control = window.child_window(title="确认另存为", control_type="Window")
        btn_yes = win_control.child_window(title="是(Y)", control_type="Button")
        btn_yes.click_input()


# 功能：浏览文件夹选择文件
# 参数：window 含有browse按钮的窗口
# name：需要选择的文件的名字
def browse_list(window, name):
    # 获取浏览的按钮并点击
    btn_browse = window.child_window(title="Browse", control_type="Button")
    btn_browse.click_input()
    # 获取弹出的文件选择窗口
    win_brow = window.child_window(title="Please choose a file", control_type="Window")
    item = win_brow.child_window(title=name, control_type="ListItem")
    item.click_input()


# 功能：以控件中心为起点，滚动
def mouse_scroll(control, distance):
    rect = control.rectangle()
    cx = int((rect.left + rect.right) / 2)
    cy = int((rect.top + rect.bottom) / 2)
    mouse.scroll(coords=(cx, cy), wheel_dist=distance)


# 功能：导出obj文件
def export_obj(window):
    time.sleep(1)
    # 获取菜单中导出文件的按钮
    btn_export = window.child_window(title="Export", control_type="Button")
    btn_export.click_input()
    # 获取弹出菜单
    stackview = window.child_window(class_name="QQuickStackView")
    # stackview.print_control_identifiers()
    time.sleep(1)
    # 获取并点击OBJ单元
    box_obj = stackview.child_window(title="OBJ", control_type="Text")
    box_obj.click_input()
    # 获取并点击最后的导出文件按钮
    btn_export1 = stackview.child_window(title="Export", control_type="Button")
    btn_export1.click_input()
    # 获取弹出的窗口，并点击其中的保存按钮
    win_brow = window.child_window(title="Please choose a file", control_type="Window")
    btn_save = win_brow.child_window(title="保存(S)", control_type="Button")
    btn_save.click_input()
    # 如果有弹窗就替换掉已经存在的转换后格式文件
    controlwindow_is_exist(win_brow)
    # 处理文件崩溃的情况
    # 如果当前的窗口不存在，则文件进入了崩溃的状态，此时返回false
    if not window.exists():
        print("文件发生崩溃")
        return -1
    else:
        # 添加对于文件导出成功与否的判断
        exported = window.child_window(title="Export completed.", control_type="Edit")
        for i in range(2):
            while not exported.exists():
                time.sleep(0.5)
                exported = window.child_window(title="Export completed.", control_type="Edit")
            print("文件导出成功！")
            break
        else:
            print("文件导出失败！")
        return -1


# 功能：导出ifc文件
def export_ifc(window):
    time.sleep(1)
    # 获取菜单中导出文件的按钮
    btn_export = window.child_window(title="Export", control_type="Button")
    btn_export.click_input()
    # 获取弹出菜单
    stackview = window.child_window(class_name="QQuickStackView")
    time.sleep(1)
    box_obj = stackview.child_window(title="IFC", control_type="Text")
    box_obj.click_input()
    # 获取并点击最后的导出文件按钮
    btn_export1 = stackview.child_window(title="Export", control_type="Button")
    btn_export1.click_input()
    # 获取弹出的窗口，并点击其中的保存按钮
    win_brow = window.child_window(title="Please choose a file", control_type="Window")
    btn_save = win_brow.child_window(title="保存(S)", control_type="Button")
    btn_save.click_input()
    # 如果有弹窗就替换掉已经存在的转换后格式文件
    controlwindow_is_exist(win_brow)
    # 处理文件崩溃的情况
    # 如果当前的窗口不存在，则文件进入了崩溃的状态，此时返回false
    if not window.exists():
        print("文件发生崩溃")
        return -1
    else:
        # 添加对于文件导出成功与否的判断
        exported = window.child_window(title="Export completed.", control_type="Edit")
        for i in range(2):
            while not exported.exists():
                time.sleep(0.5)
                exported = window.child_window(title="Export completed.", control_type="Edit")
            print("文件导出成功！")
            break
        else:
            print("文件导出失败！")
        return -1


# --------------------框架函数---------------------------
def test():
    # 获取应用
    app = get_app()
    # 获取主窗口
    win_main = app.window(class_name='ExchangerGui_MainWindow_QMLTYPE_259')
    btn_browse = win_main.child_window(title="Browse", control_type="Button")
    if btn_browse.exists():
        # browse_list(win_main, name="conrod.jt")
        browse_list(win_main, name="conrod.jt")
        # 导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("文件导入成功！")
        else:
            print("文件导入失败！")
    else:
        # 如果没有，就通过侧栏的菜单选择import按钮后再浏览目标文件
        btn_import = win_main.child_window(title="Import", control_type="Button")
        btn_import.click_input()
        # 获取弹出菜单
        stackview = win_main.child_window(class_name="QQuickStackView")
        # 选择指定文件
        browse_list(win_main, name="kk35-7.stp")
        # 获取并点击下方的导入文件按钮
        btn_import1 = stackview.child_window(title="Import", control_type="Button")
        btn_import1.click_input()
        # 对导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("文件导入成功！")
        else:
            print("文件导入失败！")
    # 导出obj格式的文件
    export_obj(win_main)


def autotest():
    # 获取应用
    app = get_app()
    # 获取主窗口
    win_main = app.window(class_name='ExchangerGui_MainWindow_QMLTYPE_259')
    btn_browse = win_main.child_window(title="Browse", control_type="Button")
    if btn_browse.exists():
        browse_list(win_main, name="conrod.jt")
        # 导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("文件导入成功！")
        else:
            print("文件导入失败！")
    else:
        # 如果没有，就通过侧栏的菜单选择import按钮后再浏览目标文件
        btn_import = win_main.child_window(title="Import", control_type="Button")
        btn_import.click_input()
        # 获取弹出菜单
        stackview = win_main.child_window(class_name="QQuickStackView")
        browse_list(win_main, name="kk35-7.stp")
        # 获取并点击下方的导入文件按钮
        btn_import1 = stackview.child_window(title="Import", control_type="Button")
        btn_import1.click_input()
        # 对导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("文件导入成功！")
        else:
            print("文件导入失败！")
    export_obj(win_main)


# ----------------------------------------------------------------

# 功能：判断是否有目标id的文件，如果有，就选中并显示选中的文件名，如果没有，就滚动
def get_id(window, id):
    # 解决某个id出现在底部的滚动条后面的问题
    bar_box = window.child_window(title="水平滚动条", auto_id="HorizontalScrollBar", control_type="ScrollBar")
    rect = bar_box.rectangle()
    rect_h = (rect.top + rect.bottom) / 2
    # 目标文件
    item = window.child_window(auto_id=str(id), control_type="ListItem")
    while not item.exists():
        mouse_scroll(control=window, distance=-5)
        item = window.child_window(auto_id=str(id), control_type="ListItem")
    rect_i = item.rectangle()
    recti_h = (rect_i.top + rect_i.bottom) / 2

    if abs(rect_h - recti_h) <= 3:
        mouse_scroll(control=window, distance=-5)

    # 解决第一页中出现的59号id问题:特殊情况单独处理
    if id == 59:
        for i in range(3):
            mouse_scroll(control=window, distance=-5)

    item.click_input()
    name = item.window_text()
    print("正在导入的文件为：%s" % name)

    # 获取并点击文件选择窗口中的打开按钮
    btn_open = window.child_window(title="打开(O)", auto_id="1", control_type="Button")
    btn_open.click_input()


# 需要实现自动化IFC的函数
# 参数i是需要获取的某个id值
def testifc(i):
    # 获取应用
    app = get_app()
    # 获取主窗口
    win_main = app.window(class_name='ExchangerGui_MainWindow_QMLTYPE_259')
    time.sleep(3)
    btn_browse = win_main.child_window(title="Browse", control_type="Button")
    # range(1, 475)
    if btn_browse.exists():
        # 点击浏览的按钮
        btn_browse.click_input()
        # 获取弹出的文件选择窗口
        win_brow = win_main.child_window(title="Please choose a file", control_type="Window")
        # 根据获取指定auto_ID的窗口
        get_id(window=win_brow, id=i)
        # 导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("第%d个文件导入成功！" % i)

            # wait_for/wait_for_not:
            #   * 'exists' means that the window is a valid handle
            #   * 'visible' means that the window is not hidden
            #   * 'enabled' means that the window is not disabled
            #   * 'ready' means that the window is visible and enabled
            #   * 'active' means that the window is active
            # timeout: 设置超时的时间，如果在n秒后没有进入指定状态，则退出等待
            # retry_interval:timeout内重试时间，How long to sleep between each retry.
            # eg: dlg.wait('ready')

            # 设置等待完成标签提示出现的时间，最大等待时间为3
            win_main.child_window(title="Display completed.", control_type="Edit").wait('exists', timeout=3)
            export_ifc(win_main)
            return 0
        else:
            print("第%d个文件导入失败！" % i)
            return -1
    else:
        # 如果没有，就通过侧栏的菜单选择import按钮后再浏览目标文件
        btn_import = win_main.child_window(title="Import", control_type="Button")
        btn_import.click_input()
        # 获取弹出菜单
        stackview = win_main.child_window(class_name="QQuickStackView")
        btn_browse = stackview.child_window(title="Browse", control_type="Button")
        btn_browse.click_input()
        # 获取弹出的文件选择窗口
        win_brow = win_main.child_window(title="Please choose a file", control_type="Window")
        get_id(window=win_brow, id=i)
        # 获取并点击下方的导入文件按钮
        btn_import1 = stackview.child_window(title="Import", control_type="Button")
        btn_import1.click_input()

        # 对导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("第%d个文件导入成功！" % i)

            win_main.child_window(title="Display completed.", control_type="Edit").wait('exists', timeout=3)
            # 问题：在执行该判定的时候程序完成按钮已经消失
            export_ifc(win_main)
            return 0
        else:
            print("第%d个文件导入失败！" % i)
            return -1


# 需要实现自动化IFC的函数-循环的方式实现
def autoifc(i):
    # 获取应用
    app = get_app()
    # 获取主窗口
    win_main = app.window(class_name='ExchangerGui_MainWindow_QMLTYPE_259')
    time.sleep(3)
    btn_browse = win_main.child_window(title="Browse", control_type="Button")
    # range(1, 475)
    if btn_browse.exists():
        # 点击浏览的按钮
        btn_browse.click_input()
        # 获取弹出的文件选择窗口
        win_brow = win_main.child_window(title="Please choose a file", control_type="Window")
        # 根据获取指定auto_ID的窗口
        get_id(window=win_brow, id=i)
        # 导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("第%d个文件导入成功！" % i)
            # 设置等待完成标签提示出现的时间，最大等待时间为3
            win_main.child_window(title="Display completed.", control_type="Edit").wait('exists', timeout=3)
            export_ifc(win_main)
            return 0
        else:
            print("第%d个文件导入失败！" % i)
            return -1
    else:
        # 如果没有，就通过侧栏的菜单选择import按钮后再浏览目标文件
        btn_import = win_main.child_window(title="Import", control_type="Button")
        btn_import.click_input()
        # 获取弹出菜单
        stackview = win_main.child_window(class_name="QQuickStackView")
        btn_browse = stackview.child_window(title="Browse", control_type="Button")
        btn_browse.click_input()
        # 获取弹出的文件选择窗口
        win_brow = win_main.child_window(title="Please choose a file", control_type="Window")
        get_id(window=win_brow, id=i)
        # 获取并点击下方的导入文件按钮
        btn_import1 = stackview.child_window(title="Import", control_type="Button")
        btn_import1.click_input()

        # 对导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("第%d个文件导入成功！" % i)
            win_main.child_window(title="Display completed.", control_type="Edit").wait('exists', timeout=3)
            # 问题：在执行该判定的时候程序完成按钮已经消失
            export_ifc(win_main)
            return 0
        else:
            print("第%d个文件导入失败！" % i)
            return -1


# 需要实现自动化OBJ的函数
def autoobj(i):
    # 获取应用
    app = get_app()
    # 获取主窗口
    win_main = app.window(class_name='ExchangerGui_MainWindow_QMLTYPE_259')
    time.sleep(3)
    btn_browse = win_main.child_window(title="Browse", control_type="Button")
    if btn_browse.exists():
        # 点击浏览的按钮
        btn_browse.click_input()
        # 获取弹出的文件选择窗口
        win_brow = win_main.child_window(title="Please choose a file", control_type="Window")
        # 根据获取指定auto_ID的窗口
        get_id(window=win_brow, id=i)
        # 导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("第%d个文件导入成功！" % i)
            # 设置等待完成标签提示出现的时间，最大等待时间为3
            win_main.child_window(title="Display completed.", control_type="Edit").wait('exists', timeout=3)
            export_obj(win_main)
            return 0
        else:
            print("第%d个文件导入失败！" % i)
            return -1
    else:
        # 如果没有，就通过侧栏的菜单选择import按钮后再浏览目标文件
        btn_import = win_main.child_window(title="Import", control_type="Button")
        btn_import.click_input()
        # 获取弹出菜单
        stackview = win_main.child_window(class_name="QQuickStackView")
        btn_browse = stackview.child_window(title="Browse", control_type="Button")
        btn_browse.click_input()
        # 获取弹出的文件选择窗口
        win_brow = win_main.child_window(title="Please choose a file", control_type="Window")
        get_id(window=win_brow, id=i)
        # 获取并点击下方的导入文件按钮
        btn_import1 = stackview.child_window(title="Import", control_type="Button")
        btn_import1.click_input()

        # 对导入成功的判断
        if win_main.child_window(title="Import completed.", control_type="Edit").exists():
            print("第%d个文件导入成功！" % i)
            win_main.child_window(title="Display completed.", control_type="Edit").wait('exists', timeout=3)
            # 问题：在执行该判定的时候程序完成按钮已经消失
            export_obj(win_main)
            return 0
        else:
            print("第%d个文件导入失败！" % i)
            return -1
