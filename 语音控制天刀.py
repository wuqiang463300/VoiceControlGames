
# -*- coding: UTF-8 -*-
from win32com.client import constants
import os
import win32com.client
import pythoncom
import  win32con
import  win32api
import time
class SpeechRecognition:
    def __init__(self, wordsToAdd):
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.listener = win32com.client.Dispatch("SAPI.SpSharedRecognizer")
        self.context = self.listener.CreateRecoContext()
        self.grammar = self.context.CreateGrammar()
        self.grammar.DictationSetState(0)
        self.wordsRule = self.grammar.Rules.Add("wordsRule", constants.SRATopLevel + constants.SRADynamic, 0)
        self.wordsRule.Clear()
        [self.wordsRule.InitialState.AddWordTransition(None, word) for word in wordsToAdd]
        self.grammar.Rules.Commit()
        self.grammar.CmdSetRuleState("wordsRule", 1)
        self.grammar.Rules.Commit()
        self.eventHandler = ContextEvents(self.context)
        self.say("Started successfully")
    def say(self, phrase):
        self.speaker.Speak(phrase)
class ContextEvents(win32com.client.getevents("SAPI.SpSharedRecoContext")):
    def OnRecognition(self, StreamNumber, StreamPosition, RecognitionType, Result):
        newResult = win32com.client.Dispatch(Result)
        print("你在说 ", newResult.PhraseInfo.GetText())
        speechstr=newResult.PhraseInfo.GetText()
        if speechstr=="前进":
                win32api.keybd_event(87, 0, 0, 0)
                time.sleep(10)
                win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="后退":
            win32api.keybd_event(83, 0, 0, 0)
            time.sleep(10)
            win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr == "左转":
            win32api.keybd_event(65,0, 0, 0)
            time.sleep(10)
            win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr == "右转":
            win32api.keybd_event(68, 0, 0, 0)
            time.sleep(10)
            win32api.keybd_event(68, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="跳起来":
            win32api.keybd_event(32, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(32, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="骑马":
            win32api.keybd_event(87, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(86, 0, 0, 0)
            time.sleep(10)
            win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="前翻":
            win32api.keybd_event(16, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(87, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(16, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="后翻":
            win32api.keybd_event(16, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(83, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(16,0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="左翻":
            win32api.keybd_event(16, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(65, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(16, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="右翻":
            win32api.keybd_event(16, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(68, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(68, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(16, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="打坐":
            win32api.keybd_event(72, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(72, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="1":
            win32api.keybd_event(49, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(49, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="2":
            win32api.keybd_event(50, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(50, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="3":
            win32api.keybd_event(51, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(51, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="4":
            win32api.keybd_event(52, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(52, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="5":
            win32api.keybd_event(53, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(53, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr == "6":
            win32api.keybd_event(54, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(54, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="7":
            win32api.keybd_event(55, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(55, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="8":
            win32api.keybd_event(56, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(56, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="9":
            win32api.keybd_event(57, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(57, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="q":
            win32api.keybd_event(81, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(81, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="e":
            win32api.keybd_event(69, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(69, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="r":
            win32api.keybd_event(82, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(82, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="t":
            win32api.keybd_event(84, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(84, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="g":
            win32api.keybd_event(71, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(71, 0, win32con.KEYEVENTF_KEYUP, 0)

if __name__ == '__main__':
    wordsToAdd = ["前进", "后退","左转","右转","前翻","后翻","左翻","右翻","跳起来","骑马","打坐", "1","2","3","4","5","6","7","8","9","q","e","r","t","g"]
    speechReco = SpeechRecognition(wordsToAdd)
    while True:
        pythoncom.PumpWaitingMessages()
