using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.Serialization;
using System.Web;

/// <summary>
/// POCO classes in the style of the Graph SDK for entities that are still in Beta and thus not covered by the SDK itself.
/// </summary>
namespace GraphModelsExtension
{
    public class Plan
    {
        public string id{ get; set; }
    }

    public class Bucket
    {
        public string id { get; set; }
        public string planId { get; set; }
        public string name { get; set; }
    }

    [DataContract()]
    public class PlannerTask
    {
        [DataMember(Name = "id")]
        public string id { get; set; }

        public EntityTagHeaderValue etag { get; set; }

        [DataMember(Name = "title")]
        public string title { get; set; }

        [DataMember(Name = "assignedTo")]
        public string assignedTo { get; set; }

        [DataMember(Name = "assignedBy")]
        public string assignedBy { get; set; }

        [DataMember(Name = "assignedDateTime")]
        public DateTimeOffset? assignedDateTime { get; set; }

        [DataMember(Name = "dueDateTime")]
        public DateTimeOffset? dueDateTime { get; set; }

        [DataMember(Name = "planId")]
        public string planId { get; set; }

        [DataMember(Name = "bucketId")]
        public string bucketId { get; set; }

        [DataMember(Name = "percentComplete")]
        public int? percentComplete { get; set; }
    }

    public class TaskDetails
    {
        public string description { get; set; }
        public string previewType { get; set; }
    }

    public class Notebook
    {
        public bool isDefault{ get; set; }
        public bool isShared { get; set; }
        public string sectionsUrl { get; set; }
        public string sectionGroupsUrl { get; set; }

        public string oneNoteWebUrl { get; set; }
        
        public string id { get; set; }
        public string name { get; set; }
    }
    public class Section
    {
        public string id { get; set; }
        public string name { get; set; }
        public string pagesUrl { get; set; }
        
    }

    public class BaseItem 
    {
        public string id { get; set; }
        public DateTimeOffset? createdDateTime { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string webUrl { get; set; }
        public DateTimeOffset? lastModifiedDateTime { get; set; }
    }

    public class Site : BaseItem
    {
        public SiteCollection siteCollection { get; set;}
        public Guid? siteCollectionId { get; set; }
        public Guid? siteId { get; set; }
    }

    [DataContract]
    public class List : BaseItem
    {
        [DataMember(Name = "list")]
        public ListInfo listInfo { get; set; }

        [DataMember(Name = "items")]
        public ListItem[] items { get; set; }
    }

    public class ListItem : BaseItem
    {
        public Int32? listItemId { get; set; }
        public JObject columnSet { get; set; }

    }

    public class SiteCollection
    {
        public string hostname { get; set; }
    }

    public class ListInfo
    {
        public bool? hidden { get; set; }
        public string template { get; set; }
    }

    /// <summary>
    /// Class to enable easy deserialization of a Graph collection that returns an object with a value property containing the collection.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ValueWrapper<T> where T : class
    {
        public T value { get; set; }
    }
}