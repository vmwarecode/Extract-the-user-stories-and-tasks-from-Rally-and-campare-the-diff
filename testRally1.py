#-------------------------------------------------------------------------------
# Name:        module 1
# Purpose:
#
# Author:      yans
#
# Created:     29/05/2015
# Copyright:   (c) yans 2015
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os
from time import *
import ConfigParser
import pyral
import requests
import sys
from pyral import Rally, rallyWorkset
from pyExcelerator import *


def setFontStyle(fontname):
    font0 = Font()
    font0.name = fontname
    font0.bold = True

    style0 = XFStyle()
    style0.font = font0
    return style0

#Create WorkSheet and Columns for Projects
def CreateProjectsWorkSheet(configFile, section, excelFileName):
    if os.path.exists(configFile):
      projectsValues = getConfigValues(configFile,section)
    ws = w.add_sheet(projectsValues['worksheetname'])
    style0 = setFontStyle('Times New Roman')
    ws.write(0,0,projectsValues['column1'],style0)
    ws.write(0,1,projectsValues['column2'],style0)
    w.save(excelFileName)

#Create WorkSheet and Columns for User Stories
def CreateUserStoriesWorkSheet(configFile, section, excelFileName):
    if os.path.exists(configFile):
      userstoriesValues = getConfigValues(configFile,section)
    ws = w.add_sheet(userstoriesValues['worksheetname'])
    style0 = setFontStyle('Times New Roman')
    ws.write(0,0,userstoriesValues['column1'],style0)
    ws.write(0,1,userstoriesValues['column2'],style0)
    ws.write(0,2,userstoriesValues['column3'],style0)
    ws.write(0,3,userstoriesValues['column4'],style0)
    ws.write(0,4,userstoriesValues['column5'],style0)
    ws.write(0,5,userstoriesValues['column6'],style0)
    ws.write(0,6,userstoriesValues['column7'],style0)
    ws.write(0,7,userstoriesValues['column8'],style0)
    ws.write(0,8,userstoriesValues['column9'],style0)
    ws.write(0,9,userstoriesValues['column10'],style0)
    ws.write(0,10,userstoriesValues['column11'],style0)
    ws.write(0,11,userstoriesValues['column12'],style0)
    ws.write(0,12,userstoriesValues['column13'],style0)
    ws.write(0,13,userstoriesValues['column14'],style0)
    ws.write(0,14,userstoriesValues['column15'],style0)
    ws.write(0,15,userstoriesValues['column16'],style0)
    w.save(excelFileName)

#Create WorkSheet and Columns for Tasks
def CreateTasksWorkSheet(configFile, section, excelFileName):
    if os.path.exists(configFile):
      userTasksValues = getConfigValues(configFile,section)
    ws = w.add_sheet(userTasksValues['worksheetname'])
    style0 = setFontStyle('Times New Roman')
    ws.write(0,0,userTasksValues['column1'],style0)
    ws.write(0,1,userTasksValues['column2'],style0)
    ws.write(0,2,userTasksValues['column3'],style0)
    ws.write(0,3,userTasksValues['column4'],style0)
    ws.write(0,4,userTasksValues['column5'],style0)
    ws.write(0,5,userTasksValues['column6'],style0)
    ws.write(0,6,userTasksValues['column7'],style0)
    ws.write(0,7,userTasksValues['column8'],style0)
    ws.write(0,8,userTasksValues['column9'],style0)
    ws.write(0,9,userTasksValues['column10'],style0)
    ws.write(0,10,userTasksValues['column11'],style0)
    ws.write(0,11,userTasksValues['column12'],style0)
    ws.write(0,12,userTasksValues['column13'],style0)
    ws.write(0,13,userTasksValues['column14'],style0)
    ws.write(0,14,userTasksValues['column15'],style0)
    ws.write(0,15,userTasksValues['column16'],style0)
    w.save(excelFileName)

#Create WorkSheet and Columns for Releases
def CreateReleasesWorkSheet(configFile, section, excelFileName):
    if os.path.exists(configFile):
      releasesValues = getConfigValues(configFile,section)
    ws = w.add_sheet(releasesValues['worksheetname'])
    style0 = setFontStyle('Times New Roman')
    ws.write(0,0,releasesValues['column1'],style0)
    ws.write(0,1,releasesValues['column2'],style0)
    ws.write(0,2,releasesValues['column3'],style0)
    ws.write(0,3,releasesValues['column4'],style0)
    ws.write(0,4,releasesValues['column5'],style0)
    ws.write(0,5,releasesValues['column6'],style0)
    ws.write(0,6,releasesValues['column7'],style0)
    ws.write(0,7,releasesValues['column8'],style0)
    ws.write(0,8,releasesValues['column9'],style0)
    ws.write(0,9,releasesValues['column10'],style0)
    w.save(excelFileName)


#Get the information for Projects from Rally
def GetProjectsInfo(projects):
    global projectIndex
    projectIndex = 0
    for proj in projects:
      projectID = proj.oid
      projectName = proj.Name
      projectIndex += 1
      ws1 = w.get_sheet(0)
      ws1.write(projectIndex, 0, projectID)
      ws1.write(projectIndex, 1, projectName)



#Get the information for User Stories filtering by Projects and Interation Name from Rally
def GetUserStoriesByIterationName(project,iterationname):
    global storyIndex
    storyIndex = 0
    global taskIndex
    taskIndex = 0
    response = rally.get('UserStory', fetch=True, query=False, project=project)
    if not response.errors:
        for story in response:
            if story.Iteration != None:
                if (story.Iteration.Name == iterationname):
                    storyID = story.oid
                    storyFormattedID = story.FormattedID
                    storyName = story.Name
                    storyScheduleState = story.ScheduleState
                    storyBlocked = story.Blocked
                    storyPlanEstimate = str(story.PlanEstimate)
                    storyOwner = ''
                    if story.Owner != None and story.Owner.Name != None:
                       storyOwner= story.Owner.Name
                    storyDefects = ''
                    if (story.Defects != None and len(story.Defects)>0):
                       storyDefects = str(len(story.Defects))
                    storyCreationDate = story.CreationDate
                    storyLastUpdateDate = story.LastUpdateDate
                    storyParent = ''
                    if story.Parent != None:
                       storyParent = story.Parent.Name
                    storyItertation = iterationname
                    storyproject = project
                    storyRelease = ''
                    if story.Release != None:
                       storyRelease = story.Release.Name
                    storyDescription = story.Description
                    storyNotes = story.Notes

                    storyIndex += 1
                    ws2 = w.get_sheet(1)
                    ws2.write(storyIndex, 0, storyID)
                    ws2.write(storyIndex, 1, storyFormattedID)
                    ws2.write(storyIndex, 2, storyName)
                    ws2.write(storyIndex, 3, storyScheduleState)
                    ws2.write(storyIndex, 4, storyBlocked)
                    ws2.write(storyIndex, 5, storyPlanEstimate)
                    ws2.write(storyIndex, 6, storyOwner)
                    ws2.write(storyIndex, 7, storyDefects)
                    ws2.write(storyIndex, 8, storyCreationDate)
                    ws2.write(storyIndex, 9, storyLastUpdateDate)
                    ws2.write(storyIndex, 10, storyParent)
                    ws2.write(storyIndex, 11, storyItertation)
                    ws2.write(storyIndex, 12, storyproject)
                    ws2.write(storyIndex, 13, storyRelease)
                    ws2.write(storyIndex, 14, storyDescription)
                    ws2.write(storyIndex, 15, storyNotes)

                    if story.Tasks != None and len(story.Tasks)>0:
                       GetTasksByUserStory(story)


#Get the information for Tasks filtering by User Story from Rally
def GetTasksByUserStory(story):
    global taskIndex

    for task in story.Tasks:
        taskID = task.oid
        taskFormattedID = task.FormattedID
        taskName = task.Name
        taskBlocked = task.Blocked
        taskItertation = task.Iteration.Name
        taskState = task.State
        taskCreationDate = task.CreationDate
        taskLastUpdateDate = task.LastUpdateDate
        taskEstimate = str(task.Estimate)
        taskTimeSpent = task.TimeSpent
        taskToDo = str(task.ToDo)
        taskVersionId = task.VersionId
        taskOwner = ''
        if task.Owner != None:
           taskOwner = task.Owner.Name
        taskDescription = task.Description
        taskNotes = task.Notes
        taskUserStoryID = story.FormattedID
        taskUserStoryName = story.Name

        taskIndex += 1
        ws3 = w.get_sheet(2)
        ws3.write(taskIndex, 0, taskID)
        ws3.write(taskIndex, 1, taskFormattedID)
        ws3.write(taskIndex, 2, taskName)
        ws3.write(taskIndex, 3, taskBlocked)
        ws3.write(taskIndex, 4, taskItertation)
        ws3.write(taskIndex, 5, taskState)
        ws3.write(taskIndex, 6, taskCreationDate)
        ws3.write(taskIndex, 7, taskLastUpdateDate)
        ws3.write(taskIndex, 8, taskEstimate)
        ws3.write(taskIndex, 9, taskTimeSpent)
        ws3.write(taskIndex, 10, taskToDo)
        ws3.write(taskIndex, 11, taskVersionId)
        ws3.write(taskIndex, 12, taskOwner)
        ws3.write(taskIndex, 13, taskDescription)
        ws3.write(taskIndex, 14, taskNotes)
        ws3.write(taskIndex, 15, taskUserStoryID)
        ws3.write(taskIndex, 16, taskUserStoryName)


#Get the information for Releases from Rally
def GetReleasesInfo():
    global releaseIndex
    releaseIndex = 0
    response = rally.get('Release', fetch=True, order="ReleaseDate")
    if not response.errors:
        for release in response:
          releaseID = release.oid
          releaseName = release.Name
          releaseCreationDate = release.CreationDate
          releaseStartDate = release.ReleaseStartDate
          releaseReleaseDate = release.ReleaseDate
          releaseState = release.State
          releaseNotes = release.Notes
          releaseVersion = ''
          if release.Version != None:
             releaseVersion = release.Version
          releaseVersionId = release.VersionId
          releaseProjectName = release.Project.Name
          releaseIndex += 1
          ws4 = w.get_sheet(3)
          ws4.write(releaseIndex, 0, releaseID)
          ws4.write(releaseIndex, 1, releaseName)
          ws4.write(releaseIndex, 2, releaseCreationDate)
          ws4.write(releaseIndex, 3, releaseStartDate)
          ws4.write(releaseIndex, 4, releaseReleaseDate)
          ws4.write(releaseIndex, 5, releaseState)
          ws4.write(releaseIndex, 6, releaseNotes)
          ws4.write(releaseIndex, 7, releaseVersion)
          ws4.write(releaseIndex, 8, releaseVersionId)
          ws4.write(releaseIndex, 9, releaseProjectName)


#Login Rally
def Login(configFile,serverLoginSection):
    global rally
    if os.path.exists(configFile):
      loginValues = getConfigValues(configFile,serverLoginSection)
      serverUrl = loginValues['serverurl']
      username = loginValues['username']
      password = loginValues['password']
      #apikey = loginValues['apikey']
      workspace = loginValues['workspace']
      project = loginValues['project']
      iterationname = loginValues['iterationname']

      rally = Rally(serverUrl, username, password)
      #rally.enableLogging('mypyral.log')
      projects = rally.getProjects(workspace=workspace)

      GetProjectsInfo(projects)
      GetUserStoriesByIterationName(project,iterationname)
##      GetReleasesInfo()


#Get item value from configuration file
def getConfigValues(filename, section):
   cp = ConfigParser.ConfigParser()
   cp.read(filename)
   vals = {}
   vals.update(cp.items(section))
   return vals


if __name__ == '__main__':
    t0 = time()
    print "\nstart: %s" % ctime(t0)
    w = Workbook()
    file = 'c:Projects_for_Astro_WEM_Release_S20.xls'
    CreateProjectsWorkSheet('config_Rally.ini', 'ProjectsWorkSheet', file)
    CreateUserStoriesWorkSheet('config_Rally.ini', 'UserStoriesWorkSheet', file)
    CreateTasksWorkSheet('config_Rally.ini', 'TasksWorkSheet', file)
##    CreateReleasesWorkSheet('config_Rally.ini', 'ReleasesWorkSheet', file)
    projectIndex = 0
    storyIndex = 0
    taskIndex = 0
    releaseIndex = 0
    Login('config_Rally.ini','ServerLogin')
    w.save(file)
    t1 = time() - t0
    print "\nsince starting elapsed %.2f s" % (t1)