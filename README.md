# Taipei DevDays 2018 add-ins-workshop

## Contents

- [Extend the functionality of the Office Apps with Add-ins (Word, Excel, PowerPoint, OneNote) workshop](#extend-the-functionality-of-the-office-apps-with-add-ins-word-excel-powerpoint-onenote-workshop)
- [What does the add-in do?](#what-does-the-add-in-do)
- [Prerequisites](#prerequisites)
- [How do I get started?](#how-do-i-get-started)

## Extend the functionality of the Office Apps with Add-ins (Word, Excel, PowerPoint, OneNote) workshop

In this repo there is a sample Office add-in and tutorial modules that guide the user to add Word and Excel specific code to complete the functionality.

## What the add-in does

_Proseware Tasker_ is a collaborative tool for teams that share authoring responsibility for Word, Excel and PowerPoint documents between members of the team or even different teams.

You can create, assign, and manage tasks right inside the document edit session using a simple, but powerful list format.

![Task creation screenshot](screenshot-createtask.png)

## Prerequisites

- Office account tenant for your team
- Visual Studio (Community is fine)
- Git command line tools
- Web browser (Chrome or Edge are fine)

## Get started

1. Start by cloning this whole repository to your local system.

    `git clone https://github.com/InteropEvents/add-ins-workshop.git`

1. Get your tenant and _Planner_ ready for the add-in by following the steps in [the setup document](setup/setup.md).

    **Tip**: Avoid path length issues with packages by cloning the repository low in the file hierarchy, like `c:\myrepos` or something with a similarly short path length.

## Follow the tutorial

1. Now you are ready to follow the tutorial. Start with [Module 1](module1/module1.md), which walks you through adding Word-specific code to the task creation process in _Proseware Tasker_.

1. Complete the tutorial by following [Module 2](module2/module2.md). This adds Excel-specific code.