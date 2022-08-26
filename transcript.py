from argparse import RawDescriptionHelpFormatter
from openpyxl.styles import Alignment
from openpyxl import load_workbook
from openpyxl import Workbook
import wget
import os
from functools import cmp_to_key


def interval(val):
    if val == 100:
        return 0
    elif val >= 90:
        return 1
    elif val >= 80:
        return 2
    elif val >= 70:
        return 3
    elif val >= 60:
        return 4
    else:
        return 5

def get_transcript_template():
    if os.path.exists('template.xlsx'):
        return load_workbook('template.xlsx')
    else:
        wget


def cmp_score(lhs, rhs):
    return rhs.weightedScoreSum - lhs.weightedScoreSum

def get_top_ten_score(studentTranscripts):
    ret = []
    for i in range(10):
        ret.append(studentTranscripts[i].weightedScoreAvg)
    return ret

def get_intervals(studentTranscripts):
    ret = [[0 for x in range(7)] for y in range(6)] 

    for student in studentTranscripts:
        for i in range(0, 7):
            ret[interval(student.studentScore[i])][i] += 1
    return ret



class Student():
    def __init__(self, studentId, rawClassScore):
        self.studentId = studentId
        self.dataRow = 0
        self.studentData = []
        self.studentScore = []
        self.name = ''
        self.transcript = None
        self.weightedScoreSum = 0.0
        self.weightedScoreAvg = 0.0
        self.rank = 0
        self.diff = 0

        for i in range(1, 50):
            if rawClassScore.cell(i, 1).value == self.studentId:
                self.dataRow = i

        for cell in rawClassScore[str(self.dataRow)]:
            self.studentData.append(cell.value)

        self.name = self.studentData[1]
        self.diff = self.studentData[12]
        self.get_score()
        
    def get_score(self):
        weights = [5, 3 ,4, 3, 1, 1, 1]
        for i in range(2, 9):
            self.studentScore.append(self.studentData[i])
        
        self.weightedScoreSum = 0.0
        for i in range(7):
            self.weightedScoreSum += weights[i] * self.studentScore[i]
        
        self.weightedScoreAvg = self.weightedScoreSum / 18

        self.studentScore.append(self.weightedScoreSum)
        self.studentScore.append(self.weightedScoreAvg)

    def generate_transcript(self, topTenScore, intervals):
        self.transcript = get_transcript_template()
        self.transcript.active['B3'] = self.name
        self.transcript.active['A4'] = self.studentId
        self.transcript.active['L4'] = self.rank
        self.transcript.active['M4'] = self.diff
        for i in range(9):
            self.transcript.active.cell(4, i+3).value = self.studentScore[i]
        
        for i in range(6):
            for j in range(7):
                self.transcript.active.cell(i+6, j+3).value = intervals[i][j]

        for i in range(10):
            self.transcript.active.cell(10+i, 13).value = topTenScore[i]

    def save_xlsx(self, path=''):
        self.transcript.save(path+'/'+str(self.studentId)+str(self.name)+'.xlsx')
            

rawClassScore = load_workbook('score.xlsx')['Sheet2']
studentIds = []

for cell in rawClassScore['a']:
    if type(cell.value) == int:
        studentIds.append(cell.value)
    
    elif type(cell.value) == None:
        break

studentTranscripts = []
topTenScore = []

for studentId in studentIds:
    studentTranscripts.append(Student(studentId, rawClassScore))

studentTranscripts = sorted(studentTranscripts, key=cmp_to_key(cmp_score))

for i in range(0, len(studentTranscripts)):
    studentTranscripts[i].rank = i+1

topTenScore = get_top_ten_score(studentTranscripts)
intervals = get_intervals(studentTranscripts)

if not os.path.exists('transcripts'):
    os.mkdir('transcripts')

for st in studentTranscripts:
    st.generate_transcript(topTenScore, intervals)
    st.save_xlsx('transcripts')
