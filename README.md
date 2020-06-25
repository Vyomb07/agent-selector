# agent-selector
Internship program 
#Excercise - Agent selector

You are given the following data for agents 
agent
→is_available
→available_since (the time since the agent is available)
→roles (a list of roles the user has, e.g. spanish speaker, sales, support etc.) 

→When an issue comes in we need to present the issue to 1 or many agents based on an agent selection mode. An agent selection mode can be all available, least busy or random. In “all available mode” the issue is presented to all agents so they pick the issue if they want. In least busy the issue is presented to the agent that has been available for the longest. In random mode we randomly pick an agent. An issue also has one or many roles (sales/support e.g.). Issues are presented to agents only with matching roles.

→Please write a function that takes an input the list of agents with their data, agent selection mode and returns a list of agents the issue should be presented to.  




Note: The automated data file is given below, save the data.xlsx where the agent.py file is present. 
      Data.xlsx is a test data which contains the details about the agent i.e Name , Availability Status, Time of Availability, Role. 
