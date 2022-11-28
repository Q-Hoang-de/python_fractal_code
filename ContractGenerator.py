import docx2pdf
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from docx2pdf import convert
import os


class ContractGenerator:
    def __init__(self, employeeName, employeeAddress, beginDate, title, salary, salaryInWord, hourWage,
                 terminationPeriod, probation):
        self.employeeName = employeeName
        self.employeeAddress = employeeAddress
        self.beginDate = beginDate
        self.title = title
        self.salary = salary
        self.salaryInWord = salaryInWord
        self.hourWage = hourWage
        self.terminationPeriod = terminationPeriod
        self.probation = probation

    def createEmployeeFullInfo(self):
        return '%s, %s' % (self.employeeName, self.employeeAddress)

    def calculateWeeklyHour(self):
        salaryInNumber = float(str(self.salary).replace(',', '.'))
        hourWageInNumber = float(str(self.hourWage).replace(',', '.'))
        return round((salaryInNumber/hourWageInNumber/26*13), 2)

    def toPdf(self, file):
        convert(file)

    def generateContract(self):
        # some fixed variable
        employer = 'Thanh Hung Tran \nInh. das Restaurant KoKoNo, Schnaitheimer Str.9, 89520 Heidenheim'
        employer_title = '-- nachfolgend Arbeitgeber genannt --'
        employee_title = '-- nachfolgend Arbeitnehmer genannt --'

        employee = self.createEmployeeFullInfo()
        weeklyHours = self.calculateWeeklyHour()

        # create actual document
        document = Document()
        document.add_heading('Arbeitsvertrag', 0)

        document.add_paragraph("Zwischen")
        document.add_paragraph().add_run(employer).bold = True
        document.add_paragraph(employer_title)

        document.add_paragraph().add_run().add_break(WD_BREAK.LINE_CLEAR_RIGHT)
        document.add_paragraph("und")
        document.add_paragraph().add_run().add_break(WD_BREAK.LINE_CLEAR_RIGHT)

        document.add_paragraph().add_run(employee).bold = True
        document.add_paragraph(employee_title)

        document.add_paragraph().add_run().add_break(WD_BREAK.LINE_CLEAR_RIGHT)

        document.add_paragraph("wird folgender Anstellungsvertrag geschlossen:")
        document.add_paragraph().add_run().add_break(WD_BREAK.LINE_CLEAR_RIGHT)

        ###
        firstSectionHeading = document.add_heading('§ 1 Arbeitsbeginn, Tätigkeitsbereich')
        document.add_paragraph('Der Arbeitnehmer tritt am %s.' % (self.beginDate))
        document.add_paragraph('Der Arbeitnehmer wird als %s beschäftigt. ' \
                               'Der Arbeitgeber kann den dem Arbeitnehmer zugewiesenen Aufgabenbereich je nach den ' \
                               'geschäftlichen Erfordernissen ergänzen oder auch ändern. Der Arbeitnehmer verpflichtet ' \
                               'sich darüber hinaus, vorübergehend auch in anderen Betriebsstätten des Arbeitgebers tätig zu sein. ' \
                               'Der Anspruch des Arbeitnehmers auf die Gehaltszahlung nach Maßgabe des § 3 dieses Vertrages bleibt ' \
                               'hiervon unberührt.' % (
                                   self.title)).paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        ###
        secondSectionHeading = document.add_heading('§ 2 Arbeitszeit')
        secondSection = document.add_paragraph(
            'Die regelmäßige wöchentliche Arbeitszeit beträgt %s Stunden.' \
            'Die Lage der Arbeitszeit und der Pausen richtet sich nach den betrieblichen Gepflogenheiten.' \
            % (str(weeklyHours)))
        ###
        thirdSectionHeading = document.add_heading('§ 3 Vergütung')
        document.add_paragraph('Für seine Tätigkeit erhält der Arbeitnehmer ein Monatsbruttogehalt in Höhe von %s €'
                               '\n(in Worten: %s).' % (str(self.salary), self.salaryInWord))
        document.add_paragraph('Stundenlohn beträgt: %s Euro.' % (str(self.hourWage)))
        document.add_paragraph('Mit dem Gehalt sind sämtliche Ansprüche des Arbeitnehmers auf Überstunden ' \
                               'bzw. Mehrarbeit bzw. Sonn- und Feiertagsarbeit abgegolten. ' \
                               'Eine Vergütung solcher Zeiten findet im Übrigen nur statt, wenn dies im Einzelfall vom ' \
                               'Arbeitgeber schriftlich zugesagt worden ist.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        ###
        fourthSectionHeading = document.add_heading('§ 4 Erholungsurlaub')
        document.add_paragraph('Der Urlaub wird nach dem Bundesurlaubsgesetzes geregelt.')
        document.add_paragraph(
            'Über den gesetzlichen Urlaubsanspruch nach Absatz 1 hinaus hat der Arbeitnehmer einen übergesetzlichen Anspruch auf bezahlten Jahresurlaub von 0 '
            'weiteren Tagen. Der Arbeitnehmer hat daher einen Urlaubsanspruch von insgesamt 24 Arbeitstagen jährlich.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph(
            'Mit der Urlaubserteilung durch den Arbeitgeber erfüllt dieser zunächst den Anspruch des Arbeitnehmers auf Urlaub nach Absatz 1, '
            'im Anschluss daran den etwaig weitergehenden Anspruch nach Absatz 2.'
            'Zeitpunkt und Dauer des Urlaubs richten sich nach den betrieblichen Notwendigkeiten und Möglichkeiten unter Berücksichtigung '
            'der Wünsche des Arbeitnehmers').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph('Das Urlaubsjahr ist das Kalenderjahr.')

        document.add_paragraph(
            'Der Urlaub muss im laufenden Kalenderjahr gewährt und genommen werden. Eine Übertragung des Urlaubs auf das nächste Kalenderjahr ist nur statthaft, '
            'wenn dringende betriebliche oder in der Person des Arbeitnehmers liegende Gründe dies rechtfertigen. '
            'Im Fall der Übertragung muss der Urlaub innerhalb der ersten drei Monate des folgenden Kalenderjahres genommen werden. '
            'Andernfalls verfällt er mit Ablauf des 31.03. des Folgejahres, soweit nicht zwingende gesetzliche Vorgaben etwas Anderes bestimmen.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph(
            'Kann der gesetzliche Urlaub wegen der Beendigung des Arbeitsverhältnisses ganz oder teilweise nicht gewährt werden, ist er abzugelten. '
            'In Bezug auf den gesetzlichen Urlaubsanspruch besteht ein Abgeltungsanspruch auch dann, wenn die Inanspruchnahme wegen krankheitsbedingter '
            'Arbeitsunfähigkeit nicht bis zum Ende des Kalenderjahres bzw. – für den Fall der Übertragung – bis zum 31.03 des Folgejahres erfolgt ist. '
            'Dies gilt allerdings längstenfalls bis zum 31.03. des übernächsten Jahres. (Beispiel: Der Urlaubsanspruch für das Jahr 2016 verfällt auch in '
            'Fällen krankheitsbedingter Arbeitsunfähigkeit spätestens am 31.03.2018). ').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        document.add_paragraph('Eine Abgeltung des übergesetzlichen Urlaubsanspruchs ist ausgeschlossen. '
                               'Dieser Anspruch erlischt mit der Beendigung des Arbeitsverhältnisses ersatzlos. Ansprüche des Arbeitnehmers auf Urlaub sind nicht vererblich, '
                               'soweit sie nicht vor dem Tod des Arbeitnehmers entstanden sind. Ansprüche auf Urlaub und Urlaubsabgeltung, die den übergesetzlichen Urlaub betreffen, '
                               'sind in keinem Fall vererblich.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph(
            'Für die Zeiten der Elternzeit gewährt der Arbeitgeber keinen (anteiligen) Urlaub. Ein Anspruch auf Urlaubsabgeltung besteht insoweit nicht. '
            'Der Arbeitgeber macht bereits jetzt von seinem Recht nach § 17 Abs. 1 Satz 1 BEEG Gebrauch.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        document.add_paragraph(
            'Im Übrigen gelten ergänzend die Bestimmungen des Bundesurlaubsgesetzes in der jeweils geltenden Fassung.')
        ###
        fifthSectionHeading = document.add_heading('§ 5 Arbeitsverhinderung')
        document.add_paragraph('Der Arbeitnehmer verpflichtet sich, jede Arbeitsverhinderung unverzüglich, '
                               'tunlichst noch vor Dienstbeginn, dem Arbeitgeber unter Benennung der voraussichtlichen Verhinderungsdauer, '
                               'ggf. telefonisch, mitzuteilen.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph(
            'Im Krankheitsfall hat der Arbeitnehmer unverzüglich, spätestens jedoch vor Ablauf des dritten Kalendertages, '
            'dem Arbeitgeber eine ärztlich erstellte Arbeitsunfähigkeitsbescheinigung vorzulegen, aus der sich die voraussichtliche Dauer der '
            'Krankheit ergibt. Dauert die Krankheit länger an als in der ärztlich erstellten Bescheinigung angegeben, so ist der '
            'Arbeitnehmer gleichfalls zur unverzüglichen Mitteilung und Vorlage einer weiteren Bescheinigung verpflichtet.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph(
            'Im Falle der Freistellung des Arbeitnehmers zur Pflege seines erkrankten Kindes erfolgt keine Entgeltfortzahlung.')
        document.add_paragraph(
            'Im Übrigen gelten für den Krankheitsfall die jeweils maßgeblichen gesetzlichen Bestimmungen.')

        ###
        sixthSectionHeading = document.add_heading('§ 6 Einstellungsfragebogen')
        document.add_paragraph('Der als Anlage beigefügte Einstellungsfragebogen ist Bestandteil dieses Vertrages. '
                               'Der Arbeitnehmer versichert die Vollständigkeit und Richtigkeit der gemachten Angaben').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        ###
        seventhSectionHeading = document.add_heading('§ 7 Nebenbeschäftigung/Wettbewerbsverbot')
        document.add_paragraph(
            'Der Arbeitnehmer hat seine gesamte Arbeitskraft ausschließlich dem Arbeitgeber zur Verfügung zu stellen. '
            'Eine Nebenbeschäftigung während des Arbeitsverhältnisses darf nur mit vorheriger schriftlicher Zustimmung des Arbeitgebers übernommen werden.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        document.add_paragraph(
            'Der Arbeitnehmer verpflichtet sich, nach seinem Ausscheiden nicht aktiv um Kunden des Arbeitgebers zu werben.')

        ###
        document.add_heading('§ 8 Kündigungsfristen')
        document.add_paragraph(
            'Das Arbeitsverhältnis wird %s eingegangen. Die ersten sechs Monate, also die Zeit bis zum %s gelten als Probezeit. '
            'Während dieser Zeit kann das Arbeitsverhältnis von beiden Seiten mit einer Frist von zwei Wochen (§ 622 Abs. (3) BGB) gekündigt werden.' % (
                self.terminationPeriod, self.probation)).paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        document.add_paragraph(
            'Nach Ablauf der Probezeit gelten die gesetzlichen Kündigungsfristen. Verlängerte Kündigungsfristen aufgrund '
            'verlängerter Betriebszugehörigkeiten gelten für beide Vertragsparteien.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        document.add_paragraph('Jede Kündigung hat schriftlich zu erfolgen.')
        document.add_paragraph('Der Arbeitgeber ist berechtigt, den Arbeitnehmer nach Ausspruch einer Kündigung '
                               'unter Fortzahlung der Vergütung und Anrechnung auf Resturlaubsansprüche von der Arbeitsleistung freizustellen.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        ###
        document.add_heading('§ 9 Kündigung vor Dienstantritt/Vertragsstrafe')
        document.add_paragraph(
            'Eine Kündigung vor Dienstantritt ist ausgeschlossen. Kündigt der Arbeitnehmer vor Dienstantritt oder hält er im '
            'Falle einer Kündigung die für ihn geltende Kündigungsfrist nicht ein, gilt eine Vertragsstrafe in Höhe eines Bruttomonatsverdienstes '
            '(vgl. § 3 dieses Vertrages) als vereinbart. Weitergehende Ansprüche des Arbeitgebers auf Schadenersatz und/oder '
            'Unterlassung bleiben hiervon unberührt.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        ###
        document.add_heading('§ 10 Krankheiten, Behinderung')
        document.add_paragraph(
            'Der Arbeitnehmer versichert, dass er nach seiner Kenntnis derzeit an keiner Krankheit leidet, die ihn an der ordnungsgemäßen '
            'Wahrnehmung seiner in diesem Vertrag bestehenden Pflichten hindert.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph(
            'Der Arbeitnehmer bestätigt ferner, dass er kein behinderter Mensch oder ein diesem Gleichgestellter im Sinne des SGB IX ist. '
            'Der Arbeitnehmer verpflichtet sich, dem Arbeitgeber unverzüglich Mitteilung zu machen, wenn er einen Antrag auf '
            'Anerkennung als behinderter Mensch bzw. Gleichgestellter stellt.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        ###
        document.add_heading('§ 11 Verfallklausel')
        document.add_paragraph(
            'Sämtliche Ansprüche aus dem Arbeitsverhältnis sind von beiden Vertragsparteien innerhalb einer Frist von sechs Monaten nach Fälligkeit der jeweils anderen Vertragspartei '
            'schriftlich gegenüber geltend zu machen. Erfolgt diese Geltendmachung nicht, gelten die Ansprüche als verfallen.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph(
            'Werden die nach Abs. (1) rechtzeitig geltend gemachten Ansprüche von der Gegenseite abgelehnt oder erklärt sich die Gegenseite nicht innerhalb von einem '
            'Monat nach der Geltendmachung des Anspruches, so verfällt dieser, wenn er nicht innerhalb von zwei Monaten nach der Ablehnung '
            'oder dem Fristablauf gerichtlich anhängig gemacht wird.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        ###
        document.add_heading('§ 12 Schlussbestimmungen')
        document.add_paragraph(
            'Mündliche Nebenabreden sind nicht getroffen worden. Änderungen und/oder Ergänzungen dieser '
            'Vereinbarung bedürfen zu ihrer Rechtswirksamkeit der Schriftform. Dies gilt auch für ein Abweichen vom Schriftformerfordernis selbst.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        document.add_paragraph(
            'Sollte eine Bestimmung dieses Vertrages unwirksam sein oder werden, so wird die Wirksamkeit der übrigen Bestimmungen '
            'davon nicht berührt. Anstelle der unwirksamen Bestimmung werden die Parteien eine solche Bestimmung treffen, '
            'die dem mit der unwirksamen Bestimmung beabsichtigten Zweck am nächsten kommt. Dies gilt auch für die Ausfüllung eventueller Vertragslücken.').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        ###
        ### Date, Location
        document.add_paragraph().add_run().add_break(WD_BREAK.LINE_CLEAR_RIGHT)
        document.add_paragraph('Ort, Datum ').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        document.add_paragraph(
            '.........................................................................................').paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        # row[1].text = 'Ort, Datum'
        ### Signature
        document.add_paragraph().add_run().add_break(WD_BREAK.LINE_CLEAR_RIGHT)
        data = (('.........................................................................................',
                 '.........................................................................................'),
                ('Unterschrift Arbeitsnehmer', 'Unterschrift Arbeitsgeber'))

        tableForSignature = document.add_table(rows=2, cols=2)
        tableForSignature.alignment = WD_TABLE_ALIGNMENT.CENTER
        for field, fieldName in data:
            row = tableForSignature.add_row().cells
            row[0].text = field
            row[1].text = fieldName
        ###

        ## Formatting document
        for para in document.paragraphs:
            para.paragraph_format.line_spacing = 1.15
            para.paragraph_format.first_line_indent = Inches(-0.25)

        document.styles['Normal']
        filename = self.getFileName(self.employeeName, self.salary)
        document.save(filename)
        self.toPdf(filename)
        self.deleteDocFile(filename)

    def getFileName(self, employeeName, salary):
        return 'Contracts/' + employeeName + '_' + str(salary) + '.docx'

    def deleteDocFile(self, file):
        os.remove(file)
