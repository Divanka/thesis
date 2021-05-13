#Author-
#Description-

import adsk.core, adsk.fusion, adsk.cam, traceback
import urllib.request
import urllib.error
import urllib.parse
import pathlib
import json

handlers = [] #события все
app = adsk.core.Application.get()
ui  = app.userInterface
global execution
global num

def run(context):
    ui = None

    #Количество муфт в проекте при запуске надстройки
    global num
    num = 0
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        workSpace = ui.workspaces.itemById('FusionSolidEnvironment')
        tbPanels = workSpace.toolbarPanels
        tbPanel = tbPanels.itemById('NewPanel')
        if tbPanel:
            tbPanel.deleteMe()
        tbPanel = tbPanels.add('NewPanel', 'Муфты', 'SelectPanel', False)

        cmdDef = ui.commandDefinitions.itemById('NewCommand')
        if cmdDef:
            cmdDef.deleteMe()

            
        cmdDef = ui.commandDefinitions.addButtonDefinition('NewCommand', 'Создать муфту', 'Ввод параметров для соединительных муфт','.//resource')
        tbPanel.controls.addCommand(cmdDef)

        ui.messageBox('Добавлена надстройка для соединительных муфт')

        #Событие по нажатию на надстройку
        sampleCommandCreated = SampleCommandCreatedEventHandler()
        cmdDef.commandCreated.add(sampleCommandCreated)
        handlers.append(sampleCommandCreated)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class SampleCommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler): #Создание событий по нажатию на надстроку
    def __init__(self):
        super().__init__()
    def notify(self, args):
        eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
        cmd = eventArgs.command

        # Вызывает создание диалогового окна
        onExecute = SampleCommandExecuteHandler()
        cmd.execute.add(onExecute)
        handlers.append(onExecute)

        #Вызывает забирание детали из сервера, кидает в папку, из папки в проект, очищает папку
        onModel = SampleCommandCrateModelHandler()
        cmd.execute.add(onModel)
        handlers.append(onModel)



class SampleCommandCrateModelHandler(adsk.core.CommandEventHandler): #Забирает сборку с облака
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface

            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            rootComp = design.rootComponent
            importManager = app.importManager

            # url = "http://195.133.144.86:4200//Half-coupling1.f3d"
            # url = "http://195.133.144.86:4200//AssemblyCoupling.f3d"
            url = "http://195.133.144.86:4200//MOVPReadyAssembly.f3d"


            doc = app.activeDocument
            if not doc.dataFile:
                ui.messageBox('Please save the document once.')
                return adsk.fusion.Occurrence.cast(None)

            # создание папки для файла
            folder = pathlib.Path(r'C:\Temp')
            if not folder.exists():
                folder.mkdir()
            #задание имени для файла по имени на сервере
            parsed = urllib.parse.urlparse(url)
            filename = parsed.path.split('//')[-1]
            dlFile = folder / filename

            # провка расширения файла с сервера
            if dlFile.suffix != '.f3d':
                ui.messageBox('F3D File Only')
                return adsk.fusion.Occurrence.cast(None)

            # если файл если, то удалить
            if dlFile.is_file():
                dlFile.unlink()

            # загрузка файла
            try:
                data = urllib.request.urlopen(url).read()
                with open(str(dlFile), mode="wb") as f:
                    f.write(data)
            except:
                ui.messageBox(f'File not found in URL\n{url}')
                return adsk.fusion.Occurrence.cast(None)

            # импортирование
            archiveOptionsNew = importManager.createFusionArchiveImportOptions(str(dlFile))
            importManager.importToTarget(archiveOptionsNew, rootComp)
            
            # удаление скаченного файла
            if dlFile.is_file():
                dlFile.unlink()

        except:
            ui.messageBox('Command executed failed: {}'.format(traceback.format_exc()))


class SampleCommandExecuteHandler(adsk.core.CommandEventHandler): #создание диалогового окна и ожидание действия в HTML
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface

            palette = ui.palettes.itemById('myExport')

            if not palette:
                palette = ui.palettes.add('myExport', 'Муфты', 'index.html', True, True, True, 600, 400)
                palette.dockingState = adsk.core.PaletteDockingStates.PaletteDockStateRight

                #должен быть запуск параметризации после нажатия на кнопку
                onHTMLEvent = MyHTMLEventHandler()
                palette.incomingFromHTML.add(onHTMLEvent)   
                handlers.append(onHTMLEvent)

            else: 
                palette.isVisible = True
        except:
            ui.messageBox('Command executed failed: {}'.format(traceback.format_exc()))


class MyHTMLEventHandler(adsk.core.HTMLEventHandler): #По нажатию кнопки проверяется нажатое и меняются параметры у сборки
    def __init__(self):
        super().__init__()
    def notify(self, args):
        global execution
        global num
        global numSp1, numSp2, numSp3, numSp4
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface
            design = app.activeProduct

            htmlArgs = adsk.core.HTMLEventArgs.cast(args)            
            data = json.loads(htmlArgs.data)

            if data['action']=='clickOk':
                ui.messageBox('я работаю')
                #после подтверждения выбора увеличиваю будующий индекс при создании еще одной муфты
                num += 1
                palette = ui.palettes.itemById('myExport')
                if palette :
                    palette.isVisible = False 
                
            if data['action']=='click': #если был клик, то забираем аргументы и проверяем выбранное исполнение. Параметризация сборки
                args = data['arguments']
                if args['shaftend'] == 'long' and args['shafthole'] == 'cylindrical': #если длинный и цилиндр
                    execution = '1'
                    
                    #Параметризация по заданному ключу в диалоговом окне
                    #Полумуфта 
                    DHC1, dHC1, D3HC1, dFHC1, l1HC1, l3HC1, bHC1, t2HC1, D1HC1, r1HC1, zHC1 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

                    dataHC1 = [("9 mm", "71 mm", "22 mm", "17 mm", "20 mm", "12 mm", "3 mm", "1.4 mm", "45 mm", "0.15 mm", "3"),
                                ("10 mm", "71 mm", "22 mm", "17 mm", "23 mm", "12 mm", "4 mm", "1.8 mm", "45 mm", "0.15 mm", "3"),
                                ("11 mm", "71 mm", "22 mm", "17 mm", "23 mm", "12 mm", "4 mm", "1.8 mm", "45 mm", "0.15 mm", "3"),
                                ("12 mm", "75 mm", "25 mm", "17 mm", "30 mm", "12 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                                ("14 mm", "75 mm", "25 mm", "17 mm", "30 mm", "12 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                                ("16 mm", "75 mm", "30 mm", "17 mm", "40 mm", "12 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                                ("18 mm", "90 mm", "32 mm", "20 mm", "40 mm", "20 mm", "6 mm", "2.8 mm", "62 mm", "0.2 mm", "4"),
                                ("20 mm", "100 mm", "38 mm", "20 mm", "50 mm", "20 mm", "6 mm", "2.8 mm", "72 mm", "0.2 mm", "6"),
                                ("22 mm", "100 mm", "38 mm", "20 mm", "50 mm", "20 mm", "7 mm", "3.3 mm", "72 mm", "0.2 mm", "6"),
                                ("25 mm", "120 mm", "50 mm", "28 mm", "60 mm", "32 mm", "7 mm", "3.3 mm", "84 mm", "0.2 mm", "6"),
                                ("28 mm", "120 mm", "50 mm", "28 mm", "60 mm", "32 mm", "8 mm", "3.3 mm", "84 mm", "0.2 mm", "6"),
                                ("32 mm", "140 mm", "67 mm", "36 mm", "80 mm", "35 mm", "10 mm", "3.3 mm", "105 mm", "0.3 mm", "6"),
                                ("36 mm", "140 mm", "67 mm", "36 mm", "80 mm", "35 mm", "10 mm", "3.3 mm", "105 mm", "0.3 mm", "6"),
                                ("40 mm", "140 mm", "75 mm", "36 mm", "110 mm", "35 mm", "12 mm", "3.8 mm", "105 mm", "0.3 mm", "6"),
                                ("45 mm", "140 mm", "75 mm", "36 mm", "110 mm", "35 mm", "14 mm", "3.8 mm", "105 mm", "0.3 mm", "6"),
                                ("50 mm", "190 mm", "95 mm", "36 mm", "110 mm", "40 mm", "16 mm", "4.3 mm", "140 mm", "0.35 mm", "8"),
                                ("56 mm", "190 mm", "95 mm", "36 mm", "110 mm", "40 mm", "16 mm", "4.3 mm", "140 mm", "0.35 mm", "8"),
                                ("63 mm", "220 mm", "120 mm", "36 mm", "140 mm", "40 mm", "18 mm", "4.4 mm", "170 mm", "0.35 mm", "10"),
                                ("71 mm", "250 mm", "130 mm", "48 mm", "140 mm", "48 mm", "20 mm", "4.9 mm", "190 mm", "0.5 mm", "10"),
                                ("80 mm", "250 mm", "140 mm", "48 mm", "170 mm", "48 mm", "22 mm", "5.4 mm", "190 mm", "0.5 mm", "10"),
                                ("90 mm", "250 mm", "150 mm", "48 mm", "170 mm", "48 mm", "24 mm", "5.4 mm", "190 mm", "0.5 mm", "10"),
                                ("100 mm", "400 mm", "220 mm", "75 mm", "210 mm", "75 mm", "28 mm", "6.4 mm", "300 mm", "0.5 mm", "10"),
                                ("110 mm", "400 mm", "220 mm", "75 mm", "210 mm", "75 mm", "32 mm", "7.4 mm", "300 mm", "0.5 mm", "10"),
                                ("125 mm", "400 mm", "220 mm", "75 mm", "210 mm", "75 mm", "32 mm", "7.4 mm", "300 mm", "0.5 mm", "10"),
                                ("140 mm", "500 mm", "250 mm", "90 mm", "250 mm", "90 mm", "36 mm", "8.4 mm", "350 mm", "0.9 mm", "12"),
                                ("160 mm", "500 mm", "250 mm", "90 mm", "300 mm", "90 mm", "40 mm", "9.4 mm", "350 mm", "0.9 mm", "12")]

                    key = args['key'] + ' mm'

                    for i in range(0, len(dataHC1)):
                        if dataHC1[i][0] == key:
                            dHC1 = dataHC1[i][0]
                            DHC1 = dataHC1[i][1]
                            D3HC1 = dataHC1[i][2]
                            dFHC1 = dataHC1[i][3]
                            l1HC1 = dataHC1[i][4]
                            l3HC1 = dataHC1[i][5]
                            bHC1 = dataHC1[i][6]
                            t2HC1 = dataHC1[i][7]
                            D1HC1 = dataHC1[i][8]
                            r1HC1 = dataHC1[i][9]
                            zHC1 = dataHC1[i][10]
                            break

                    if num == 0: #если это первое нажатие на надстройку
                        dHC1Param = design.userParameters.itemByName('dHC1')
                        dHC1Param.expression = dHC1 
                        dHC1ForConParam = design.userParameters.itemByName('dHC1ForCon')
                        dHC1ForConParam.expression = dHC1 
                        DHC1Param = design.userParameters.itemByName('DHC1')
                        DHC1Param.expression = DHC1
                        D3HC1Param = design.userParameters.itemByName('D3HC1')
                        D3HC1Param.expression = D3HC1
                        dFHC1Param = design.userParameters.itemByName('dFHC1')
                        dFHC1Param.expression = dFHC1
                        l1HC1Param = design.userParameters.itemByName('l1HC1')
                        l1HC1Param.expression = l1HC1
                        l3HC1Param = design.userParameters.itemByName('l3HC1')
                        l3HC1Param.expression = l3HC1
                        bHC1Param = design.userParameters.itemByName('bHC1')
                        bHC1Param.expression = bHC1
                        bHC1ForConParam = design.userParameters.itemByName('bHC1ForCon')
                        bHC1ForConParam.expression = bHC1
                        t2HC1Param = design.userParameters.itemByName('t2HC1')
                        t2HC1Param.expression = t2HC1
                        t2HC1ForConParam = design.userParameters.itemByName('t2HC1ForCon')
                        t2HC1ForConParam.expression = t2HC1
                        D1HC1Param = design.userParameters.itemByName('D1HC1')
                        D1HC1Param.expression = D1HC1
                        r1HC1Param = design.userParameters.itemByName('r1HC1')
                        r1HC1Param.expression = r1HC1
                        zHC1Param = design.userParameters.itemByName('zHC1')
                        zHC1Param.expression = zHC1

                    elif num != 0:   #если это не первое нажатие на надстройку  
                        dHC1Param = design.userParameters.itemByName('dHC1_' + str(num))
                        dHC1Param.expression = dHC1   
                        dHC1ForConParam = design.userParameters.itemByName('dHC1ForCon_' + str(num))
                        dHC1ForConParam.expression = dHC1 
                        DHC1Param = design.userParameters.itemByName('DHC1_' + str(num))
                        DHC1Param.expression = DHC1
                        D3HC1Param = design.userParameters.itemByName('D3HC1_' + str(num))
                        D3HC1Param.expression = D3HC1
                        dFHC1Param = design.userParameters.itemByName('dFHC1_' + str(num))
                        dFHC1Param.expression = dFHC1
                        l1HC1Param = design.userParameters.itemByName('l1HC1_' + str(num))
                        l1HC1Param.expression = l1HC1
                        l3HC1Param = design.userParameters.itemByName('l3HC1_' + str(num))
                        l3HC1Param.expression = l3HC1
                        bHC1Param = design.userParameters.itemByName('bHC1_' + str(num))
                        bHC1Param.expression = bHC1
                        bHC1ForConParam = design.userParameters.itemByName('bHC1ForCon_' + str(num))
                        bHC1ForConParam.expression = bHC1
                        t2HC1Param = design.userParameters.itemByName('t2HC1_' + str(num))
                        t2HC1Param.expression = t2HC1
                        t2HC1ForConParam = design.userParameters.itemByName('t2HC1ForCon_' + str(num))
                        t2HC1ForConParam.expression = t2HC1
                        D1HC1Param = design.userParameters.itemByName('D1HC1_' + str(num))
                        D1HC1Param.expression = D1HC1
                        r1HC1Param = design.userParameters.itemByName('r1HC1_' + str(num))
                        r1HC1Param.expression = r1HC1
                        zHC1Param = design.userParameters.itemByName('zHC1_' + str(num))
                        zHC1Param.expression = zHC1

                    #Полумуфта
                    DHC2, dHC2, D3HC2, dFHC2, l1HC2, l2HC2, bHC2, t2HC2, D1HC2, r1HC2, zHC2 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

                    dataHC2 = [("9 mm", "71 mm", "22 mm", "8 mm", "20 mm", "9 mm", "3 mm", "1.4 mm", "45 mm", "0.15 mm", "3"),
                            ("10 mm", "71 mm", "22 mm", "8 mm", "23 mm", "9 mm", "4 mm", "1.8 mm", "45 mm", "0.15 mm", "3"),
                            ("11 mm", "71 mm", "22 mm", "8 mm", "23 mm", "9 mm", "4 mm", "1.8 mm", "45 mm", "0.15 mm", "3"),
                            ("12 mm", "75 mm", "25 mm", "8 mm", "30 mm", "9 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                            ("14 mm", "75 mm", "25 mm", "8 mm", "30 mm", "9 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                            ("16 mm", "75 mm", "30 mm", "10 mm", "40 mm", "9 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                            ("18 mm", "90 mm", "32 mm", "10 mm", "40 mm", "16 mm", "6 mm", "2.8 mm", "62 mm", "0.2 mm", "4"),
                            ("20 mm", "100 mm", "38 mm", "10 mm", "50 mm", "16 mm", "6 mm", "2.8 mm", "72 mm", "0.2 mm", "6"),
                            ("22 mm", "100 mm", "38 mm", "10 mm", "50 mm", "16 mm", "7 mm", "3.3 mm", "72 mm", "0.2 mm", "6"),
                            ("25 mm", "120 mm", "50 mm", "14 mm", "60 mm", "18 mm", "7 mm", "3.3 mm", "84 mm", "0.2 mm", "6"),
                            ("28 mm", "120 mm", "50 mm", "14 mm", "60 mm", "18 mm", "8 mm", "3.3 mm", "84 mm", "0.2 mm", "6"),
                            ("32 mm", "140 mm", "67 mm", "14 mm", "80 mm", "22 mm", "10 mm", "3.3 mm", "105 mm", "0.3 mm", "6"),
                            ("36 mm", "140 mm", "67 mm", "14 mm", "80 mm", "22 mm", "10 mm", "3.3 mm", "105 mm", "0.3 mm", "6"),
                            ("40 mm", "140 mm", "75 mm", "14 mm", "110 mm", "22 mm", "12 mm", "3.8 mm", "105 mm", "0.3 mm", "6"),
                            ("45 mm", "140 mm", "75 mm", "14 mm", "110 mm", "22 mm", "14 mm", "3.8 mm", "105 mm", "0.3 mm", "6"),
                            ("50 mm", "190 mm", "95 mm", "18 mm", "110 mm", "24 mm", "16 mm", "4.3 mm", "140 mm", "0.35 mm", "8"),
                            ("56 mm", "190 mm", "95 mm", "18 mm", "110 mm", "24 mm", "16 mm", "4.3 mm", "140 mm", "0.35 mm", "8"),
                            ("63 mm", "220 mm", "120 mm", "18 mm", "140 mm", "24 mm", "18 mm", "4.4 mm", "170 mm", "0.35 mm", "10"),
                            ("71 mm", "250 mm", "130 mm", "24 mm", "140 mm", "30 mm", "20 mm", "4.9 mm", "190 mm", "0.5 mm", "10"),
                            ("80 mm", "250 mm", "140 mm", "24 mm", "170 mm", "30 mm", "22 mm", "5.4 mm", "190 mm", "0.5 mm", "10"),
                            ("90 mm", "250 mm", "150 mm", "24 mm", "170 mm", "30 mm", "24 mm", "5.4 mm", "190 mm", "0.5 mm", "10"),
                            ("100 mm", "400 mm", "220 mm", "38 mm", "210 mm", "48 mm", "28 mm", "6.4 mm", "300 mm", "0.5 mm", "10"),
                            ("110 mm", "400 mm", "220 mm", "38 mm", "210 mm", "48 mm", "32 mm", "7.4 mm", "300 mm", "0.5 mm", "10"),
                            ("125 mm", "400 mm", "220 mm", "38 mm", "210 mm", "48 mm", "32 mm", "7.4 mm", "300 mm", "0.5 mm", "10"),
                            ("140 mm", "500 mm", "250 mm", "45 mm", "250 mm", "70 mm", "36 mm", "8.4 mm", "350 mm", "0.9 mm", "12"),
                            ("160 mm", "500 mm", "250 mm", "45 mm", "300 mm", "70 mm", "40 mm", "9.4 mm", "350 mm", "0.9 mm", "12")]


                    for i in range(0, len(dataHC2)):
                        if dataHC2[i][0] == key:
                            dHC2 = dataHC2[i][0]
                            DHC2 = dataHC2[i][1]
                            D3HC2 = dataHC2[i][2]
                            dFHC2 = dataHC2[i][3]
                            l1HC2 = dataHC2[i][4]
                            l2HC2 = dataHC2[i][5]
                            bHC2 = dataHC2[i][6]
                            t2HC2 = dataHC2[i][7]
                            D1HC2 = dataHC2[i][8]
                            r1HC2 = dataHC2[i][9]
                            zHC2 = dataHC2[i][10]
                            break
                    if num == 0:
                        dHC2Param = design.userParameters.itemByName('dHC2')
                        dHC2Param.expression = dHC2
                        dHC2ForConParam = design.userParameters.itemByName('dHC2ForCon')
                        dHC2ForConParam.expression = dHC2
                        DHC2Param = design.userParameters.itemByName('DHC2')
                        DHC2Param.expression = DHC2
                        D3HC2Param = design.userParameters.itemByName('D3HC2')
                        D3HC2Param.expression = D3HC2
                        dFHC2Param = design.userParameters.itemByName('dFHC2')
                        dFHC2Param.expression = dFHC2
                        l1HC2Param = design.userParameters.itemByName('l1HC2')
                        l1HC2Param.expression = l1HC2
                        l2HC2Param = design.userParameters.itemByName('l2HC2')
                        l2HC2Param.expression = l2HC2
                        bHC2Param = design.userParameters.itemByName('bHC2')
                        bHC2Param.expression = bHC2
                        bHC2ForConParam = design.userParameters.itemByName('bHC2ForCon')
                        bHC2ForConParam.expression = bHC2
                        t2HC2Param = design.userParameters.itemByName('t2HC2')
                        t2HC2Param.expression = t2HC2
                        t2HC2ForConParam = design.userParameters.itemByName('t2HC2ForCon')
                        t2HC2ForConParam.expression = t2HC2
                        D1HC2Param = design.userParameters.itemByName('D1HC2')
                        D1HC2Param.expression = D1HC2
                        r1HC2Param = design.userParameters.itemByName('r1HC2')
                        r1HC2Param.expression = r1HC2
                        zHC2Param = design.userParameters.itemByName('zHC2')
                        zHC2Param.expression = zHC2

                    elif num != 0:
                        dHC2Param = design.userParameters.itemByName('dHC2_' + str(num))
                        dHC2Param.expression = dHC2
                        dHC2ForConParam = design.userParameters.itemByName('dHC2ForCon_' + str(num))
                        dHC2ForConParam.expression = dHC2
                        DHC2Param = design.userParameters.itemByName('DHC2_' + str(num))
                        DHC2Param.expression = DHC2
                        D3HC2Param = design.userParameters.itemByName('D3HC2_' + str(num))
                        D3HC2Param.expression = D3HC2
                        dFHC2Param = design.userParameters.itemByName('dFHC2_' + str(num))
                        dFHC2Param.expression = dFHC2
                        l1HC2Param = design.userParameters.itemByName('l1HC2_' + str(num))
                        l1HC2Param.expression = l1HC2
                        l2HC2Param = design.userParameters.itemByName('l2HC2_' + str(num))
                        l2HC2Param.expression = l2HC2
                        bHC2Param = design.userParameters.itemByName('bHC2_' + str(num))
                        bHC2Param.expression = bHC2
                        bHC2ForConParam = design.userParameters.itemByName('bHC2ForCon_' + str(num))
                        bHC2ForConParam.expression = bHC2
                        t2HC2Param = design.userParameters.itemByName('t2HC2_' + str(num))
                        t2HC2Param.expression = t2HC2
                        t2HC2ForConParam = design.userParameters.itemByName('t2HC2ForCon_' + str(num))
                        t2HC2ForConParam.expression = t2HC2
                        D1HC2Param = design.userParameters.itemByName('D1HC2_' + str(num))
                        D1HC2Param.expression = D1HC2
                        r1HC2Param = design.userParameters.itemByName('r1HC2_' + str(num))
                        r1HC2Param.expression = r1HC2
                        zHC2Param = design.userParameters.itemByName('zHC2_' + str(num))
                        zHC2Param.expression = zHC2

                    #Проставка (втулка распорная)
                    hSp, dSp, DSp = 0, 0, 0

                    dataSp = [('8 mm',"3 mm", "12 mm"),
                            ('10 mm',"4 mm", "14 mm"),
                            ('14 mm',"5 mm", "20 mm"),
                            ('18 mm',"6 mm", "25 mm"),
                            ('24 mm',"8 mm", "32 mm"),
                            ('38 mm',"12 mm", "46 mm"),
                            ('45 mm',"15 mm", "55 mm")]

                    keyF = dFHC2

                    for i in range(0, len(dataSp)):
                        if dataSp[i][0] == keyF:
                            dSp = dataSp[i][0]
                            hSp = dataSp[i][1]
                            DSp = dataSp[i][2]
                            break

                    if num == 0:
                        hSpParam = design.userParameters.itemByName('hSp')
                        hSpParam.expression = hSp
                        DSpParam = design.userParameters.itemByName('DSp')
                        DSpParam.expression = DSp
                        dSpParam = design.userParameters.itemByName('dSp')
                        dSpParam.expression = dSp
                    elif num != 0:
                        hSpParam = design.userParameters.itemByName('hSp_' + str(num))
                        hSpParam.expression = hSp
                        DSpParam = design.userParameters.itemByName('DSp_' + str(num))
                        DSpParam.expression = DSp
                        dSpParam = design.userParameters.itemByName('dSp_' + str(num))
                        dSpParam.expression = dSp
                        
                    #Втулки резиновые
                    dataSl = [("8 mm", "17 mm", "3 mm", "1.5 mm"),
                            ("10 mm", "20 mm", "5 mm", "2.5 mm"),
                            ("14 mm", "28 mm", "7 mm", "3.5 mm"),
                            ("18 mm", "36 mm", "9 mm", "4.5 mm"),
                            ("24 mm", "48 mm", "11 mm", "6 mm"),
                            ("38 mm", "75 mm", "18 mm", "10 mm"),
                            ("45 mm", "90 mm", "22 mm", "12 mm")]

                    dSl, DSl, h1Sl, h2Sl = 0, 0, 0, 0

                    keySl = dFHC2

                    for i in range(0, len(dataSl)):
                        if dataSl[i][0] == keySl:
                            dSl = dataSl[i][0]
                            DSl = dataSl[i][1]
                            h1Sl = dataSl[i][2]
                            h2Sl = dataSl[i][3]
                            break

                    if num == 0:
                        #Втулка 1            
                        dSlParam = design.userParameters.itemByName('dSl')
                        dSlParam.expression = dSl
                        DSlParam = design.userParameters.itemByName('DSl')
                        DSlParam.expression = DSl
                        h1SlParam = design.userParameters.itemByName('h1Sl')
                        h1SlParam.expression = h1Sl
                        h2SlParam = design.userParameters.itemByName('h2Sl')
                        h2SlParam.expression = h2Sl
                        #Втулка 2
                        dSl_1Param = design.userParameters.itemByName('dSl_1')
                        dSl_1Param.expression = dSl
                        DSl_1Param = design.userParameters.itemByName('DSl_1')
                        DSl_1Param.expression = DSl
                        h1Sl_1Param = design.userParameters.itemByName('h1Sl_1')
                        h1Sl_1Param.expression = h1Sl
                        h2Sl_1Param = design.userParameters.itemByName('h2Sl_1')
                        h2Sl_1Param.expression = h2Sl
                        #Втулка 3
                        dSl_2Param = design.userParameters.itemByName('dSl_2')
                        dSl_2Param.expression = dSl
                        DSl_2Param = design.userParameters.itemByName('DSl_2')
                        DSl_2Param.expression = DSl
                        h1Sl_2Param = design.userParameters.itemByName('h1Sl_2')
                        h1Sl_2Param.expression = h1Sl
                        h2Sl_2Param = design.userParameters.itemByName('h2Sl_2')
                        h2Sl_2Param.expression = h2Sl
                        #Втулка 4
                        dSl_3Param = design.userParameters.itemByName('dSl_3')
                        dSl_3Param.expression = dSl
                        DSl_3Param = design.userParameters.itemByName('DSl_3')
                        DSl_3Param.expression = DSl
                        h1Sl_3Param = design.userParameters.itemByName('h1Sl_3')
                        h1Sl_3Param.expression = h1Sl
                        h2Sl_3Param = design.userParameters.itemByName('h2Sl_3')
                        h2Sl_3Param.expression = h2Sl

                    elif num != 0:
                        #Втулка 1            
                        dSlParam = design.userParameters.itemByName('dSl' + '_' + str(num))
                        dSlParam.expression = dSl
                        DSlParam = design.userParameters.itemByName('DSl' + '_' + str(num))
                        DSlParam.expression = DSl
                        h1SlParam = design.userParameters.itemByName('h1Sl' + '_' + str(num))
                        h1SlParam.expression = h1Sl
                        h2SlParam = design.userParameters.itemByName('h2Sl' + '_' + str(num))
                        h2SlParam.expression = h2Sl
                        #Втулка 2
                        dSl_1Param = design.userParameters.itemByName('dSl_1' + '_' + str(num))
                        dSl_1Param.expression = dSl
                        DSl_1Param = design.userParameters.itemByName('DSl_1' + '_' + str(num))
                        DSl_1Param.expression = DSl
                        h1Sl_1Param = design.userParameters.itemByName('h1Sl_1' + '_' + str(num))
                        h1Sl_1Param.expression = h1Sl
                        h2Sl_1Param = design.userParameters.itemByName('h2Sl_1' + '_' + str(num))
                        h2Sl_1Param.expression = h2Sl
                        #Втулка 3
                        dSl_2Param = design.userParameters.itemByName('dSl_2' + '_' + str(num))
                        dSl_2Param.expression = dSl
                        DSl_2Param = design.userParameters.itemByName('DSl_2' + '_' + str(num))
                        DSl_2Param.expression = DSl
                        h1Sl_2Param = design.userParameters.itemByName('h1Sl_2' + '_' + str(num))
                        h1Sl_2Param.expression = h1Sl
                        h2Sl_2Param = design.userParameters.itemByName('h2Sl_2' + '_' + str(num))
                        h2Sl_2Param.expression = h2Sl
                        #Втулка 4
                        dSl_3Param = design.userParameters.itemByName('dSl_3' + '_' + str(num))
                        dSl_3Param.expression = dSl
                        DSl_3Param = design.userParameters.itemByName('DSl_3' + '_' + str(num))
                        DSl_3Param.expression = DSl
                        h1Sl_3Param = design.userParameters.itemByName('h1Sl_3' + '_' + str(num))
                        h1Sl_3Param.expression = h1Sl
                        h2Sl_3Param = design.userParameters.itemByName('h2Sl_3' + '_' + str(num))
                        h2Sl_3Param.expression = h2Sl

                    #Палец
                    dForNutFin, lForNutFin, lForHC2Fin, dFForHC2Fin, lForHC1Fin, dBigFin, lFin = 0, 0, 0, 0, 0, 0, 0

                    dataFin = [("6 mm", "12 mm", "9 mm", "8 mm", "15 mm", " 12 mm", "2 mm"),
                            ("8 mm", "14 mm", "16 mm", "10 mm", "19 mm", " 14 mm", "2 mm"),
                            ("10 mm", "18 mm", "18 mm", "14 mm", "33 mm", " 20 mm", "2 mm"),
                            ("12 mm", "23 mm", "24 mm", "18 mm", "42 mm", " 25 mm", "3 mm"),
                            ("16 mm", "28 mm", "30 mm", "24 mm", "52 mm", " 32 mm", "3 mm"),
                            ("24 mm", "40 mm", "48 mm", "38 mm", "84 mm", " 48 mm", "4 mm")]

                    key = dFHC2

                    for i in range(0, len(dataFin)):
                        if dataFin[i][3] == key:
                            dForNutFin = dataFin[i][0]
                            lForNutFin = dataFin[i][1]
                            lForHC2Fin = dataFin[i][2]
                            dFForHC2Fin = dataFin[i][3]
                            lForHC1Fin = dataFin[i][4]
                            dBigFin = dataFin[i][5]
                            lFin = dataFin[i][6]
                            break

                    if num == 0:
                        dForNutFinParam = design.userParameters.itemByName('dForNutFin')
                        dForNutFinParam.expression = dForNutFin
                        lForNutFinParam = design.userParameters.itemByName('lForNutFin')
                        lForNutFinParam.expression = lForNutFin
                        lForHC2FinParam = design.userParameters.itemByName('lForHC2Fin')
                        lForHC2FinParam.expression = lForHC2Fin
                        dFForHC2FinParam = design.userParameters.itemByName('dFForHC2Fin')
                        dFForHC2FinParam.expression = dFForHC2Fin
                        lForHC1FinParam = design.userParameters.itemByName('lForHC1Fin')
                        lForHC1FinParam.expression = lForHC1Fin
                        dBigFinParam = design.userParameters.itemByName('dBigFin')
                        dBigFinParam.expression = dBigFin
                        lFinParam = design.userParameters.itemByName('lFin')
                        lFinParam.expression = lFin

                    elif num != 0:
                        dForNutFinParam = design.userParameters.itemByName('dForNutFin_' + str(num))
                        dForNutFinParam.expression = dForNutFin
                        lForNutFinParam = design.userParameters.itemByName('lForNutFin_' + str(num))
                        lForNutFinParam.expression = lForNutFin
                        lForHC2FinParam = design.userParameters.itemByName('lForHC2Fin_' + str(num))
                        lForHC2FinParam.expression = lForHC2Fin
                        dFForHC2FinParam = design.userParameters.itemByName('dFForHC2Fin_' + str(num))
                        dFForHC2FinParam.expression = dFForHC2Fin
                        lForHC1FinParam = design.userParameters.itemByName('lForHC1Fin_' + str(num))
                        lForHC1FinParam.expression = lForHC1Fin
                        dBigFinParam = design.userParameters.itemByName('dBigFin_' + str(num))
                        dBigFinParam.expression = dBigFin
                        lFinParam = design.userParameters.itemByName('lFin_' + str(num))
                        lFinParam.expression = lFin


                if args['shaftend'] == 'short' and args['shafthole'] == 'cylindrical': #если короткий и цилиндр
                    execution = '2'
                    if args['key'] == '9':
                        ui.messageBox('По выбранным вами параметрам невозможно построить муфту с коротким концом валов')
                       
                    DHC1, dHC1, D3HC1, dFHC1, l1HC1, l3HC1, bHC1, t2HC1, D1HC1, r1HC1, zHC1 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

                    dataHC1 = [("10 mm", "71 mm", "22 mm", "17 mm", "20 mm", "12 mm", "4 mm", "1.8 mm", "45 mm", "0.15 mm", "3"),
                                ("11 mm", "71 mm", "22 mm", "17 mm", "20 mm", "12 mm", "4 mm", "1.8 mm", "45 mm", "0.15 mm", "3"),
                                ("12 mm", "75 mm", "25 mm", "17 mm", "25 mm", "12 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                                ("14 mm", "75 mm", "25 mm", "17 mm", "25 mm", "12 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                                ("16 mm", "75 mm", "30 mm", "17 mm", "28 mm", "12 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                                ("18 mm", "90 mm", "32 mm", "20 mm", "28 mm", "20 mm", "6 mm", "2.8 mm", "62 mm", "0.2 mm", "4"),
                                ("20 mm", "100 mm", "38 mm", "20 mm", "36 mm", "20 mm", "6 mm", "2.8 mm", "72 mm", "0.2 mm", "6"),
                                ("22 mm", "100 mm", "38 mm", "20 mm", "36 mm", "20 mm", "7 mm", "3.3 mm", "72 mm", "0.2 mm", "6"),
                                ("25 mm", "120 mm", "50 mm", "28 mm", "42 mm", "32 mm", "7 mm", "3.3 mm", "84 mm", "0.2 mm", "6"),
                                ("28 mm", "120 mm", "50 mm", "28 mm", "42 mm", "32 mm", "8 mm", "3.3 mm", "84 mm", "0.2 mm", "6"),
                                ("32 mm", "140 mm", "67 mm", "32 mm", "58 mm", "35 mm", "10 mm", "3.3 mm", "105 mm", "0.3 mm", "6"),
                                ("36 mm", "140 mm", "67 mm", "32 mm", "58 mm", "35 mm", "10 mm", "3.3 mm", "105 mm", "0.3 mm", "6"),
                                ("40 mm", "140 mm", "75 mm", "32 mm", "82 mm", "35 mm", "12 mm", "3.8 mm", "105 mm", "0.3 mm", "6"),
                                ("45 mm", "140 mm", "75 mm", "32 mm", "82 mm", "35 mm", "14 mm", "3.8 mm", "105 mm", "0.3 mm", "6"),
                                ("50 mm", "190 mm", "95 mm", "36 mm", "82 mm", "40 mm", "16 mm", "4.3 mm", "140 mm", "0.35 mm", "8"),
                                ("56 mm", "190 mm", "95 mm", "36 mm", "82 mm", "40 mm", "16 mm", "4.3 mm", "140 mm", "0.35 mm", "8"),
                                ("63 mm", "220 mm", "120 mm", "36 mm", "105 mm", "40 mm", "18 mm", "4.4 mm", "170 mm", "0.35 mm", "10"),
                                ("71 mm", "250 mm", "130 mm", "48 mm", "105 mm", "48 mm", "20 mm", "4.9 mm", "190 mm", "0.5 mm", "10"),
                                ("80 mm", "250 mm", "140 mm", "48 mm", "130 mm", "48 mm", "22 mm", "5.4 mm", "190 mm", "0.5 mm", "10"),
                                ("90 mm", "250 mm", "150 mm", "48 mm", "130 mm", "48 mm", "24 mm", "5.4 mm", "190 mm", "0.5 mm", "10"),
                                ("100 mm", "400 mm", "220 mm", "75 mm", "165 mm", "75 mm", "28 mm", "6.4 mm", "300 mm", "0.5 mm", "10"),
                                ("110 mm", "400 mm", "220 mm", "75 mm", "165 mm", "75 mm", "32 mm", "7.4 mm", "300 mm", "0.5 mm", "10"),
                                ("125 mm", "400 mm", "220 mm", "75 mm", "165 mm", "75 mm", "32 mm", "7.4 mm", "300 mm", "0.5 mm", "10"),
                                ("140 mm", "500 mm", "250 mm", "90 mm", "200 mm", "90 mm", "36 mm", "8.4 mm", "350 mm", "0.9 mm", "12"),
                                ("160 mm", "500 mm", "250 mm", "90 mm", "240 mm", "90 mm", "40 mm", "9.4 mm", "350 mm", "0.9 mm", "12")]

                    key = args['key'] + ' mm'

                    for i in range(0, len(dataHC1)):
                        if dataHC1[i][0] == key:
                            dHC1 = dataHC1[i][0]
                            DHC1 = dataHC1[i][1]
                            D3HC1 = dataHC1[i][2]
                            dFHC1 = dataHC1[i][3]
                            l1HC1 = dataHC1[i][4]
                            l3HC1 = dataHC1[i][5]
                            bHC1 = dataHC1[i][6]
                            t2HC1 = dataHC1[i][7]
                            D1HC1 = dataHC1[i][8]
                            r1HC1 = dataHC1[i][9]
                            zHC1 = dataHC1[i][10]
                            break

                    if num == 0: #если это первое нажатие на надстройку
                        dHC1Param = design.userParameters.itemByName('dHC1')
                        dHC1Param.expression = dHC1   
                        DHC1Param = design.userParameters.itemByName('DHC1')
                        DHC1Param.expression = DHC1
                        D3HC1Param = design.userParameters.itemByName('D3HC1')
                        D3HC1Param.expression = D3HC1
                        dFHC1Param = design.userParameters.itemByName('dFHC1')
                        dFHC1Param.expression = dFHC1
                        l1HC1Param = design.userParameters.itemByName('l1HC1')
                        l1HC1Param.expression = l1HC1
                        l3HC1Param = design.userParameters.itemByName('l3HC1')
                        l3HC1Param.expression = l3HC1
                        bHC1Param = design.userParameters.itemByName('bHC1')
                        bHC1Param.expression = bHC1
                        t2HC1Param = design.userParameters.itemByName('t2HC1')
                        t2HC1Param.expression = t2HC1
                        D1HC1Param = design.userParameters.itemByName('D1HC1')
                        D1HC1Param.expression = D1HC1
                        r1HC1Param = design.userParameters.itemByName('r1HC1')
                        r1HC1Param.expression = r1HC1
                        zHC1Param = design.userParameters.itemByName('zHC1')
                        zHC1Param.expression = zHC1

                    elif num != 0:   #если это не первое нажатие на надстройку  
                        dHC1Param = design.userParameters.itemByName('dHC1_' + str(num))
                        dHC1Param.expression = dHC1   
                        DHC1Param = design.userParameters.itemByName('DHC1_' + str(num))
                        DHC1Param.expression = DHC1
                        D3HC1Param = design.userParameters.itemByName('D3HC1_' + str(num))
                        D3HC1Param.expression = D3HC1
                        dFHC1Param = design.userParameters.itemByName('dFHC1_' + str(num))
                        dFHC1Param.expression = dFHC1
                        l1HC1Param = design.userParameters.itemByName('l1HC1_' + str(num))
                        l1HC1Param.expression = l1HC1
                        l3HC1Param = design.userParameters.itemByName('l3HC1_' + str(num))
                        l3HC1Param.expression = l3HC1
                        bHC1Param = design.userParameters.itemByName('bHC1_' + str(num))
                        bHC1Param.expression = bHC1
                        t2HC1Param = design.userParameters.itemByName('t2HC1_' + str(num))
                        t2HC1Param.expression = t2HC1
                        D1HC1Param = design.userParameters.itemByName('D1HC1_' + str(num))
                        D1HC1Param.expression = D1HC1
                        r1HC1Param = design.userParameters.itemByName('r1HC1_' + str(num))
                        r1HC1Param.expression = r1HC1
                        zHC1Param = design.userParameters.itemByName('zHC1_' + str(num))
                        zHC1Param.expression = zHC1

                    #Полумуфта
                    DHC2, dHC2, D3HC2, dFHC2, l1HC2, l2HC2, bHC2, t2HC2, D1HC2, r1HC2, zHC2 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

                    dataHC2 = [("10 mm", "71 mm", "22 mm", "8 mm", "20 mm", "9 mm", "4 mm", "1.8 mm", "45 mm", "0.15 mm", "3"),
                            ("11 mm", "71 mm", "22 mm", "8 mm", "20 mm", "9 mm", "4 mm", "1.8 mm", "45 mm", "0.15 mm", "3"),
                            ("12 mm", "75 mm", "25 mm", "8 mm", "25 mm", "9 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                            ("14 mm", "75 mm", "25 mm", "8 mm", "25 mm", "9 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                            ("16 mm", "75 mm", "30 mm", "10 mm", "28 mm", "9 mm", "5 mm", "2.3 mm", "50 mm", "0.2 mm", "4"),
                            ("18 mm", "90 mm", "32 mm", "10 mm", "28 mm", "16 mm", "6 mm", "2.8 mm", "62 mm", "0.2 mm", "4"),
                            ("20 mm", "100 mm", "38 mm", "10 mm", "36 mm", "16 mm", "6 mm", "2.8 mm", "72 mm", "0.2 mm", "6"),
                            ("22 mm", "100 mm", "38 mm", "10 mm", "36 mm", "16 mm", "7 mm", "3.3 mm", "72 mm", "0.2 mm", "6"),
                            ("25 mm", "120 mm", "50 mm", "14 mm", "42 mm", "18 mm", "7 mm", "3.3 mm", "84 mm", "0.2 mm", "6"),
                            ("28 mm", "120 mm", "50 mm", "14 mm", "42 mm", "18 mm", "8 mm", "3.3 mm", "84 mm", "0.2 mm", "6"),
                            ("32 mm", "140 mm", "67 mm", "14 mm", "58 mm", "22 mm", "10 mm", "3.3 mm", "105 mm", "0.3 mm", "6"),
                            ("36 mm", "140 mm", "67 mm", "14 mm", "58 mm", "22 mm", "10 mm", "3.3 mm", "105 mm", "0.3 mm", "6"),
                            ("40 mm", "140 mm", "75 mm", "14 mm", "82 mm", "22 mm", "12 mm", "3.8 mm", "105 mm", "0.3 mm", "6"),
                            ("45 mm", "140 mm", "75 mm", "14 mm", "82 mm", "22 mm", "14 mm", "3.8 mm", "105 mm", "0.3 mm", "6"),
                            ("50 mm", "190 mm", "95 mm", "18 mm", "82 mm", "24 mm", "16 mm", "4.3 mm", "140 mm", "0.35 mm", "8"),
                            ("56 mm", "190 mm", "95 mm", "18 mm", "82 mm", "24 mm", "16 mm", "4.3 mm", "140 mm", "0.35 mm", "8"),
                            ("63 mm", "220 mm", "120 mm", "18 mm", "105 mm", "24 mm", "18 mm", "4.4 mm", "170 mm", "0.35 mm", "10"),
                            ("71 mm", "250 mm", "130 mm", "24 mm", "105 mm", "30 mm", "20 mm", "4.9 mm", "190 mm", "0.5 mm", "10"),
                            ("80 mm", "250 mm", "140 mm", "24 mm", "130 mm", "30 mm", "22 mm", "5.4 mm", "190 mm", "0.5 mm", "10"),
                            ("90 mm", "250 mm", "150 mm", "24 mm", "130 mm", "30 mm", "24 mm", "5.4 mm", "190 mm", "0.5 mm", "10"),
                            ("100 mm", "400 mm", "220 mm", "38 mm", "165 mm", "48 mm", "28 mm", "6.4 mm", "300 mm", "0.5 mm", "10"),
                            ("110 mm", "400 mm", "220 mm", "38 mm", "165 mm", "48 mm", "32 mm", "7.4 mm", "300 mm", "0.5 mm", "10"),
                            ("125 mm", "400 mm", "220 mm", "38 mm", "165 mm", "48 mm", "32 mm", "7.4 mm", "300 mm", "0.5 mm", "10"),
                            ("140 mm", "500 mm", "250 mm", "45 mm", "200 mm", "70 mm", "36 mm", "8.4 mm", "350 mm", "0.9 mm", "12"),
                            ("160 mm", "500 mm", "250 mm", "45 mm", "240 mm", "70 mm", "40 mm", "9.4 mm", "350 mm", "0.9 mm", "12")]


                    for i in range(0, len(dataHC2)):
                        if dataHC2[i][0] == key:
                            dHC2 = dataHC2[i][0]
                            DHC2 = dataHC2[i][1]
                            D3HC2 = dataHC2[i][2]
                            dFHC2 = dataHC2[i][3]
                            l1HC2 = dataHC2[i][4]
                            l2HC2 = dataHC2[i][5]
                            bHC2 = dataHC2[i][6]
                            t2HC2 = dataHC2[i][7]
                            D1HC2 = dataHC2[i][8]
                            r1HC2 = dataHC2[i][9]
                            zHC2 = dataHC2[i][10]
                            break
                    if num == 0:
                        dHC2Param = design.userParameters.itemByName('dHC2')
                        dHC2Param.expression = dHC2
                        DHC2Param = design.userParameters.itemByName('DHC2')
                        DHC2Param.expression = DHC2
                        D3HC2Param = design.userParameters.itemByName('D3HC2')
                        D3HC2Param.expression = D3HC2
                        dFHC2Param = design.userParameters.itemByName('dFHC2')
                        dFHC2Param.expression = dFHC2
                        l1HC2Param = design.userParameters.itemByName('l1HC2')
                        l1HC2Param.expression = l1HC2
                        l2HC2Param = design.userParameters.itemByName('l2HC2')
                        l2HC2Param.expression = l2HC2
                        bHC2Param = design.userParameters.itemByName('bHC2')
                        bHC2Param.expression = bHC2
                        t2HC2Param = design.userParameters.itemByName('t2HC2')
                        t2HC2Param.expression = t2HC2
                        D1HC2Param = design.userParameters.itemByName('D1HC2')
                        D1HC2Param.expression = D1HC2
                        r1HC2Param = design.userParameters.itemByName('r1HC2')
                        r1HC2Param.expression = r1HC2
                        zHC2Param = design.userParameters.itemByName('zHC2')
                        zHC2Param.expression = zHC2

                    elif num != 0:
                        dHC2Param = design.userParameters.itemByName('dHC2_' + str(num))
                        dHC2Param.expression = dHC2
                        DHC2Param = design.userParameters.itemByName('DHC2_' + str(num))
                        DHC2Param.expression = DHC2
                        D3HC2Param = design.userParameters.itemByName('D3HC2_' + str(num))
                        D3HC2Param.expression = D3HC2
                        dFHC2Param = design.userParameters.itemByName('dFHC2_' + str(num))
                        dFHC2Param.expression = dFHC2
                        l1HC2Param = design.userParameters.itemByName('l1HC2_' + str(num))
                        l1HC2Param.expression = l1HC2
                        l2HC2Param = design.userParameters.itemByName('l2HC2_' + str(num))
                        l2HC2Param.expression = l2HC2
                        bHC2Param = design.userParameters.itemByName('bHC2_' + str(num))
                        bHC2Param.expression = bHC2
                        t2HC2Param = design.userParameters.itemByName('t2HC2_' + str(num))
                        t2HC2Param.expression = t2HC2
                        D1HC2Param = design.userParameters.itemByName('D1HC2_' + str(num))
                        D1HC2Param.expression = D1HC2
                        r1HC2Param = design.userParameters.itemByName('r1HC2_' + str(num))
                        r1HC2Param.expression = r1HC2
                        zHC2Param = design.userParameters.itemByName('zHC2_' + str(num))
                        zHC2Param.expression = zHC2

                    #Проставка (втулка распорная)
                    hSp, dSp, DSp = 0, 0, 0

                    dataSp = [('8 mm',"3 mm", "12 mm"),
                            ('10 mm',"4 mm", "14 mm"),
                            ('14 mm',"5 mm", "20 mm"),
                            ('18 mm',"6 mm", "25 mm"),
                            ('24 mm',"8 mm", "32 mm"),
                            ('38 mm',"12 mm", "46 mm"),
                            ('45 mm',"15 mm", "55 mm")]

                    keyF = dFHC2

                    for i in range(0, len(dataSp)):
                        if dataSp[i][0] == keyF:
                            dSp = dataSp[i][0]
                            hSp = dataSp[i][1]
                            DSp = dataSp[i][2]
                            break

                    if num == 0:
                        hSpParam = design.userParameters.itemByName('hSp')
                        hSpParam.expression = hSp
                        DSpParam = design.userParameters.itemByName('DSp')
                        DSpParam.expression = DSp
                        dSpParam = design.userParameters.itemByName('dSp')
                        dSpParam.expression = dSp
                    elif num != 0:
                        hSpParam = design.userParameters.itemByName('hSp_' + str(num))
                        hSpParam.expression = hSp
                        DSpParam = design.userParameters.itemByName('DSp_' + str(num))
                        DSpParam.expression = DSp
                        dSpParam = design.userParameters.itemByName('dSp_' + str(num))
                        dSpParam.expression = dSp
                        
                    #Втулки резиновые
                    dataSl = [("8 mm", "17 mm", "3 mm", "1.5 mm"),
                            ("10 mm", "20 mm", "5 mm", "2.5 mm"),
                            ("14 mm", "28 mm", "7 mm", "3.5 mm"),
                            ("18 mm", "36 mm", "9 mm", "4.5 mm"),
                            ("24 mm", "48 mm", "11 mm", "6 mm"),
                            ("38 mm", "75 mm", "18 mm", "10 mm"),
                            ("40 mm", "90 mm", "22 mm", "12 mm")]

                    dSl, DSl, h1Sl, h2Sl = 0, 0, 0, 0

                    keySl = dFHC2

                    for i in range(0, len(dataSl)):
                        if dataSl[i][0] == keySl:
                            dSl = dataSl[i][0]
                            DSl = dataSl[i][1]
                            h1Sl = dataSl[i][2]
                            h2Sl = dataSl[i][3]
                            break

                    if num == 0:
                        #Втулка 1            
                        dSlParam = design.userParameters.itemByName('dSl')
                        dSlParam.expression = dSl
                        DSlParam = design.userParameters.itemByName('DSl')
                        DSlParam.expression = DSl
                        h1SlParam = design.userParameters.itemByName('h1Sl')
                        h1SlParam.expression = h1Sl
                        h2SlParam = design.userParameters.itemByName('h2Sl')
                        h2SlParam.expression = h2Sl
                        #Втулка 2
                        dSl_1Param = design.userParameters.itemByName('dSl_1')
                        dSl_1Param.expression = dSl
                        DSl_1Param = design.userParameters.itemByName('DSl_1')
                        DSl_1Param.expression = DSl
                        h1Sl_1Param = design.userParameters.itemByName('h1Sl_1')
                        h1Sl_1Param.expression = h1Sl
                        h2Sl_1Param = design.userParameters.itemByName('h2Sl_1')
                        h2Sl_1Param.expression = h2Sl
                        #Втулка 3
                        dSl_2Param = design.userParameters.itemByName('dSl_2')
                        dSl_2Param.expression = dSl
                        DSl_2Param = design.userParameters.itemByName('DSl_2')
                        DSl_2Param.expression = DSl
                        h1Sl_2Param = design.userParameters.itemByName('h1Sl_2')
                        h1Sl_2Param.expression = h1Sl
                        h2Sl_2Param = design.userParameters.itemByName('h2Sl_2')
                        h2Sl_2Param.expression = h2Sl
                        #Втулка 4
                        dSl_3Param = design.userParameters.itemByName('dSl_3')
                        dSl_3Param.expression = dSl
                        DSl_3Param = design.userParameters.itemByName('DSl_3')
                        DSl_3Param.expression = DSl
                        h1Sl_3Param = design.userParameters.itemByName('h1Sl_3')
                        h1Sl_3Param.expression = h1Sl
                        h2Sl_3Param = design.userParameters.itemByName('h2Sl_3')
                        h2Sl_3Param.expression = h2Sl

                    elif num != 0:
                        #Втулка 1            
                        dSlParam = design.userParameters.itemByName('dSl' + '_' + str(num))
                        dSlParam.expression = dSl
                        DSlParam = design.userParameters.itemByName('DSl' + '_' + str(num))
                        DSlParam.expression = DSl
                        h1SlParam = design.userParameters.itemByName('h1Sl' + '_' + str(num))
                        h1SlParam.expression = h1Sl
                        h2SlParam = design.userParameters.itemByName('h2Sl' + '_' + str(num))
                        h2SlParam.expression = h2Sl
                        #Втулка 2
                        dSl_1Param = design.userParameters.itemByName('dSl_1' + '_' + str(num))
                        dSl_1Param.expression = dSl
                        DSl_1Param = design.userParameters.itemByName('DSl_1' + '_' + str(num))
                        DSl_1Param.expression = DSl
                        h1Sl_1Param = design.userParameters.itemByName('h1Sl_1' + '_' + str(num))
                        h1Sl_1Param.expression = h1Sl
                        h2Sl_1Param = design.userParameters.itemByName('h2Sl_1' + '_' + str(num))
                        h2Sl_1Param.expression = h2Sl
                        #Втулка 3
                        dSl_2Param = design.userParameters.itemByName('dSl_2' + '_' + str(num))
                        dSl_2Param.expression = dSl
                        DSl_2Param = design.userParameters.itemByName('DSl_2' + '_' + str(num))
                        DSl_2Param.expression = DSl
                        h1Sl_2Param = design.userParameters.itemByName('h1Sl_2' + '_' + str(num))
                        h1Sl_2Param.expression = h1Sl
                        h2Sl_2Param = design.userParameters.itemByName('h2Sl_2' + '_' + str(num))
                        h2Sl_2Param.expression = h2Sl
                        #Втулка 4
                        dSl_3Param = design.userParameters.itemByName('dSl_3' + '_' + str(num))
                        dSl_3Param.expression = dSl
                        DSl_3Param = design.userParameters.itemByName('DSl_3' + '_' + str(num))
                        DSl_3Param.expression = DSl
                        h1Sl_3Param = design.userParameters.itemByName('h1Sl_3' + '_' + str(num))
                        h1Sl_3Param.expression = h1Sl
                        h2Sl_3Param = design.userParameters.itemByName('h2Sl_3' + '_' + str(num))
                        h2Sl_3Param.expression = h2Sl

                    #Палец
                    dForNutFin, lForNutFin, lForHC2Fin, dFForHC2Fin, lForHC1Fin, dBigFin, lFin = 0, 0, 0, 0, 0, 0, 0

                    dataFin = [("6 mm", "12 mm", "9 mm", "8 mm", "15 mm", " 12 mm", "2 mm"),
                            ("8 mm", "14 mm", "16 mm", "10 mm", "19 mm", " 14 mm", "2 mm"),
                            ("10 mm", "18 mm", "18 mm", "14 mm", "33 mm", " 20 mm", "2 mm"),
                            ("12 mm", "23 mm", "24 mm", "18 mm", "42 mm", " 25 mm", "3 mm"),
                            ("16 mm", "28 mm", "30 mm", "24 mm", "52 mm", " 32 mm", "3 mm"),
                            ("24 mm", "40 mm", "48 mm", "38 mm", "84 mm", " 48 mm", "4 mm")]

                    key = dFHC2

                    for i in range(0, len(dataFin)):
                        if dataFin[i][3] == key:
                            dForNutFin = dataFin[i][0]
                            lForNutFin = dataFin[i][1]
                            lForHC2Fin = dataFin[i][2]
                            dFForHC2Fin = dataFin[i][3]
                            lForHC1Fin = dataFin[i][4]
                            dBigFin = dataFin[i][5]
                            lFin = dataFin[i][6]
                            break

                    if num == 0:
                        dForNutFinParam = design.userParameters.itemByName('dForNutFin')
                        dForNutFinParam.expression = dForNutFin
                        lForNutFinParam = design.userParameters.itemByName('lForNutFin')
                        lForNutFinParam.expression = lForNutFin
                        lForHC2FinParam = design.userParameters.itemByName('lForHC2Fin')
                        lForHC2FinParam.expression = lForHC2Fin
                        dFForHC2FinParam = design.userParameters.itemByName('dFForHC2Fin')
                        dFForHC2FinParam.expression = dFForHC2Fin
                        lForHC1FinParam = design.userParameters.itemByName('lForHC1Fin')
                        lForHC1FinParam.expression = lForHC1Fin
                        dBigFinParam = design.userParameters.itemByName('dBigFin')
                        dBigFinParam.expression = dBigFin
                        lFinParam = design.userParameters.itemByName('lFin')
                        lFinParam.expression = lFin

                    elif num != 0:
                        dForNutFinParam = design.userParameters.itemByName('dForNutFin_' + str(num))
                        dForNutFinParam.expression = dForNutFin
                        lForNutFinParam = design.userParameters.itemByName('lForNutFin_' + str(num))
                        lForNutFinParam.expression = lForNutFin
                        lForHC2FinParam = design.userParameters.itemByName('lForHC2Fin_' + str(num))
                        lForHC2FinParam.expression = lForHC2Fin
                        dFForHC2FinParam = design.userParameters.itemByName('dFForHC2Fin_' + str(num))
                        dFForHC2FinParam.expression = dFForHC2Fin
                        lForHC1FinParam = design.userParameters.itemByName('lForHC1Fin_' + str(num))
                        lForHC1FinParam.expression = lForHC1Fin
                        dBigFinParam = design.userParameters.itemByName('dBigFin_' + str(num))
                        dBigFinParam.expression = dBigFin
                        lFinParam = design.userParameters.itemByName('lFin_' + str(num))
                        lFinParam.expression = lFin
                    
                if args['shaftend'] == 'long' and args['shafthole'] == 'conical': #если длинный и конус
                    execution = '3'
                if args['shaftend'] == 'short' and args['shafthole'] == 'conical': #если короткий и конус
                    execution = '4'

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        ui.messageBox('Надстройка удалена')
        workSpace = ui.workspaces.itemById('FusionSolidEnvironment')
        tbPanels = workSpace.toolbarPanels
        tbPanel = tbPanels.itemById('NewPanel')
        if tbPanel:
            tbPanel.deleteMe()
        palette = ui.palettes.itemById('myExport')
        if palette:
            palette.deleteMe()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
