# cosplaydeathmatch
Array logic implemented to programatically manipulate a PowerPoint presentation through VBA macros. 
Co-developed by Andrey Zholudok, Alison Bird in 2018

## Table of Contents
1. Overview
2. Problem
3. Solutions
4. Future Design Considerations

## Overview

Cosplay Death Match is an improvisational performance event, taking 16 or 32 contestants through a 'popularity contest' in the format of a bracketed tournament. Characters are introduced in pairs and the audience votes on who goes through to the next round. The loser is elimnated until there is one winner. 
This style event can be limited to "pose" style contests (like Death Match) or "talent" style contents (like Lip Sync Battle). 

## Problem

Requested specifications: 
* Title card to announce the beginning and end of the event
* Round card to announce the start of a round and bracket
-- Must cover: Round 1, (Round 2), Quarter Final, Semi Final, Final
-- Round 2 is only for 32 contentants; skip for 16 contestants
-- Must cover: Brackets A, B, C, D
* Fight card to show the two competitors currently fighting

## Solution

* Create a maintain an array of fight results and a fight counter
* When loading a bracket or fight screen, use utilities to pull the fighter image data and replace the image data for the placeholder shapes on the slide
* When a fight is done, right the results to the fight array and increment the fight counter
* All solution code is stored in the macro-enabled PowerPoint file, but has been extracted for ease of reference

## Future Design Considerations
Microsoft PowerPoint does not respond well to dynamically manipulating the image data in shape options and transitioning between slides. The image data is frequently cached improperly. 

Future versions of this visual would be more robust if we used a "Smash Brothers Character Select" style screen, where all objects are present, but the visibility state changes - rather than the current "bracket" version which requires dynamically adjusting the visuals based on event state. 
