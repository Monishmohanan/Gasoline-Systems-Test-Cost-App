# Gasoline-Systems-Test-Cost-App

An automated Test and Cost Template Generator App for GS Components & Products

<br>

## About
The standalone application is developed for generating the test and cost template for the Gasoline Systems(GS) products
The app uses a primary database <strong>info.db</strong> for initially loading the contents
Every instance of the template generation is being recorded in the database <strong>report.db</strong> which is located in the network folder

<br>

## Workflow
- The application gets the required inputs i.e. change type, subassebmly and parts from the user
- The inputs are used to filter the tests from the test database
- The filtered tests are compared against the cost database to obtain the cost information
- The obtained information is used to generate a template that can verified, attested and shared with the product testing team

<br>

## Graphical User Interface

### Main Window

![mainwindow_data](https://user-images.githubusercontent.com/60011463/130349632-88f554b5-9dce-4914-9209-c356187e9161.PNG)

<br>
### Confirmation Window

![confirmation](https://user-images.githubusercontent.com/60011463/130350159-709a311e-e791-47a8-ae1e-7cdb0abf79e9.PNG)

<br>
### Settings Window

![settings](https://user-images.githubusercontent.com/60011463/130350009-92522e28-5d5c-41ef-a283-aed169680777.PNG)


