/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
$.getScript("../models/task.js", function() {});
$.getScript("../models/resource.js", function() {});

$(document).ready(() => {
  $("#run").click(run);
});

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
  $("#sideload-msg").hide();
  $("#app-body").show();
};

// Global variables
var tasks = [];

function run() {
  submitTasks();
  submitResourcesAssignments();
}

// Get the GUID of a task.
function getTaskGuid(index) {
  var defer = $.Deferred();
  Office.context.document.getTaskByIndexAsync(index, function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      onError(result.error);
    } else {
      defer.resolve(result.value);
    }
  });
  return defer.promise();
}

// Get the maximum index of the tasks for the current project.
function getMaxTaskIndex() {
  var defer = $.Deferred();
  Office.context.document.getMaxTaskIndexAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      onError(result.error);
    } else {
      defer.resolve(result.value);
    }
  });
  return defer.promise();
}

// Get the attribute value for a specific task.
function getTaskAttribute(taskGuid, targetField) {
  var defer = $.Deferred();
  Office.context.document.getTaskFieldAsync(taskGuid, targetField, function(
    result
  ) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      onError(result.error);
    } else {
      defer.resolve(result.value.fieldValue);
    }
  });
  return defer.promise();
}

// Get all the basic attributes values for a specific task.
function getTaskObject(index, maxIndex, resources, callback) {
  var guid, name, duration, start, finish, hasChild, parentGuid, resourceNames;
  getTaskGuid(index).then(function(taskGuid) {
    guid = taskGuid.replace(/[{}]/g, "");
    getTaskAttribute(taskGuid, Office.ProjectTaskFields.Name).then(function(
      value
    ) {
      name = value;
      getTaskAttribute(taskGuid, Office.ProjectTaskFields.Duration).then(
        function(value) {
          duration = value;
          getTaskAttribute(taskGuid, Office.ProjectTaskFields.Start).then(
            function(value) {
              start = value;
              getTaskAttribute(taskGuid, Office.ProjectTaskFields.Finish).then(
                function(value) {
                  finish = value;
                  getTaskAttribute(
                    taskGuid,
                    Office.ProjectTaskFields.Summary
                  ).then(function(value) {
                    hasChild = value;
                    getTaskAttribute(
                      taskGuid,
                      Office.ProjectTaskFields.ResourceNames
                    ).then(function(value) {
                      //console.log( value);
                      resourceNames = value;
                      // Get the parent task
                      var jump = false;
                      for (let i = index - 1; i >= 0; i--) {
                        if (tasks[i].hasChild) {
                          if (!jump) {
                            parentGuid = tasks[i].guid;
                            break;
                          } else jump = false;
                        } else if (hasChild) jump = true;
                      }
                      // Get resources Guids
                      var resourcesGuids = [];
                      var parsedResourcesNames = resourceNames.split(",");
                      $.each(parsedResourcesNames, function(
                        index,
                        resourceName
                      ) {
                        $.each(resources, function(index, resource) {
                          if (resource.name == resourceName)
                            resourcesGuids.push(resource.guid);
                        });
                      });
                      var task = new Task(
                        index,
                        guid,
                        name,
                        duration,
                        start,
                        finish,
                        hasChild,
                        parentGuid,
                        resourcesGuids
                      );
                      tasks.push(task);
                      if (index == maxIndex) {
                        var completeListOfTasks = tasks;
                        document.getElementById("spinner").style.display =
                          "none";
                        tasks = [];
                        callback(completeListOfTasks, resources);
                      }
                    });
                  });
                }
              );
            }
          );
        }
      );
    });
  });
}

function getMaxResourceIndex() {
  var defer = $.Deferred();
  Office.context.document.getMaxResourceIndexAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      onError("getMaxResourceIndex|" + JSON.stringify(result.error));
    } else {
      defer.resolve(result.value);
    }
  });
  return defer.promise();
}

function getResourceByIndex(index) {
  var defer = $.Deferred();
  Office.context.document.getResourceByIndexAsync(index, function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      onError("getResourceByIndex|" + JSON.stringify(result.error));
    } else {
      defer.resolve(result.value);
    }
  });
  return defer.promise();
}

function getResourceName(resourceGuid) {
  var defer = $.Deferred();
  Office.context.document.getResourceFieldAsync(
    resourceGuid,
    Office.ProjectResourceFields.Name,
    function(result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        onError("getResourceName|" + JSON.stringify(result.error));
      } else {
        defer.resolve(result.value.fieldValue);
      }
    }
  );
  return defer.promise();
}

function constructResources(callback) {
  getMaxResourceIndex().then(function(maxIndex) {
    var guid, name;
    var resources = [];
    for (let index = 1; index <= maxIndex; index++) {
      getResourceByIndex(index).then(function(resourceGuid) {
        getResourceName("" + resourceGuid + "").then(function(resourceName) {
          guid = resourceGuid.replace(/[{}]/g, "");
          name = resourceName;
          var resource = new Resource(index, guid, name, []);
          resources.push(resource);
          if (index == maxIndex) callback(resources);
        });
      });
    }
  });
}

function GetAllTasksAndRes(callback) {
  document.getElementById("spinner").style.display = "block";
  constructResources(function(resources) {
    getMaxTaskIndex().then(function(maxIndex) {
      for (let index = 0; index <= maxIndex; index++) {
        getTaskObject(index, maxIndex, resources, callback);
      }
    });
  });
}

// SUPPORT FUNCTIONS
function onError(error) {
  console.log("ERROR: " + error);
}

/////////////////////////////////////////// API REQUESTS ///////////////////////////////////////////
// Send tasks to API
function submitTasks() {
  GetAllTasksAndRes(function(tasks, resources) {
    // <visualization>
    $.each(tasks, function(index, task) {
      console.dir(JSON.stringify(task));
    });
    // </visualization>

    // @Anand: here you can make the AJAX request to the API
  });
}

// Send Resources Assignments to API
function submitResourcesAssignments() {
  var assignedResources = [];
  GetAllTasksAndRes(function(tasks, resources) {
    $.each(resources, function(index, resource) {
      var tempResource = resource;
      $.each(tasks, function(index, task) {
        if (task.resourceGuids.indexOf(resource.guid) > -1) {
          tempResource.tasksNames.push(task.name);
        }
      });
      assignedResources.push(tempResource);
    });

    // <visualization>
    console.dir("THE FOLLOWING ARE THE RESOURCES:");
    $.each(assignedResources, function(index, resource) {
      console.dir(JSON.stringify(resource));
    });
    // </visualization>

    // @Anand: here you can make the AJAX request to the API
  });
}
