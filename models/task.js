// Initializing a constructor function
function Task(
  id,
  guid,
  name,
  duration,
  start,
  finish,
  hasChild,
  parentGuid,
  resourceGuids
) {
  this.id = id;
  this.guid = guid;
  this.name = name;
  this.duration = duration;
  this.start = start;
  this.finish = finish;
  this.hasChild = hasChild;
  this.parentGuid = parentGuid;
  this.resourceGuids = resourceGuids;
}
