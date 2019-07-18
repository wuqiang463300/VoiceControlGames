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
        elif speechstr=="加速":
            win32api.keybd_event(87, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(86, 0, 0, 0)
            time.sleep(10)
            win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="调息":
            win32api.keybd_event(90, 0, 0, 0)
            time.sleep(9)
            win32api.keybd_event(90, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="技能1":
            win32api.keybd_event(49, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(49, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="技能2":
            win32api.keybd_event(50, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(50, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="技能3":
            win32api.keybd_event(51, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(51, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="技能4":
            win32api.keybd_event(52, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(52, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif speechstr=="技能5":
            win32api.keybd_event(53, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(53, 0, win32con.KEYEVENTF_KEYUP, 0)

if __name__ == '__main__':
    wordsToAdd = ["前进", "后退","左转","右转","跳起来","加速","调息", "技能1","技能2","技能3","技能4","技能5"]
    speechReco = SpeechRecognition(wordsToAdd)
    while True:
        pythoncom.PumpWaitingMessages()
