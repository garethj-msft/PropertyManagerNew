﻿using Microsoft.OData.ProxyExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Graph;

namespace SuiteLevelWebApp
{
    public static class PagedCollectionExtensions
    {
        //Extension methods for the IPagedCollection interfaces returned
        //by the Microsoft.Graph.GraphService

        //The GraphServiceExtension methods use these extension methods 
        //to interact with the IPagedCollection interfaces returned by the
        //Microsoft.Graph.GraphService
        //public static async Task<TElement[]> GetAllAsnyc<TElement>(
        //    this IPagedCollection<TElement> pagedCollection)
        //{
        //    var list = new List<TElement>();

        //    var collection = pagedCollection;
        //    while (true)
        //    {
        //        list.AddRange(collection.CurrentPage);
        //        if (!collection.MorePagesAvailable) break;
        //        collection = await collection.GetNextPageAsync();
        //    }
        //    return list.ToArray();
        //}

        public static async Task<Message[]> GetAllAsync(this IUserMessagesCollectionRequest pagedCollectionRq)
        {
            var list = new List<Message>();

            var collectionRequest = pagedCollectionRq;
            while (true && collectionRequest != null)
            {
                var pageList = await collectionRequest.GetAsync();
                if (pageList.CurrentPage.Count > 0)
                {
                    list.AddRange(pageList.CurrentPage);
                    collectionRequest = pageList.NextPageRequest;
                }
                else
                {
                    break;
                }
            }
            return list.ToArray();
        }
        public static async Task<DriveItem[]> GetAllAsync(this IDriveItemChildrenCollectionRequest pagedCollectionRq)
        {
            var list = new List<DriveItem>();

            var collectionRequest = pagedCollectionRq;
            while (true && collectionRequest != null)
            {
                var pageList = await collectionRequest.GetAsync();
                if (pageList.CurrentPage.Count > 0)
                {
                    list.AddRange(pageList.CurrentPage);
                    collectionRequest = pageList.NextPageRequest;
                }
                else
                {
                    break;
                }
            }
            return list.ToArray();
        }
        public static async Task<Conversation[]> GetAllAsync(this IGroupConversationsCollectionRequest pagedCollectionRq)
        {
            var list = new List<Conversation>();

            var collectionRequest = pagedCollectionRq;
            while (true && collectionRequest != null)
            {
                var pageList = await collectionRequest.GetAsync();
                if (pageList.CurrentPage.Count > 0)
                {
                    list.AddRange(pageList.CurrentPage);
                    collectionRequest = pageList.NextPageRequest;
                }
                else
                {
                    break;
                }
            }
            return list.ToArray();
        }
        public static async Task<User[]> GetAllAsync(this IGraphServiceUsersCollectionRequest pagedCollectionRq)
        {
            var list = new List<User>();

            var collectionRequest = pagedCollectionRq;
            while (true && collectionRequest != null)
            {
                var pageList = await collectionRequest.GetAsync();
                if (pageList.CurrentPage.Count > 0)
                {
                    list.AddRange(pageList.CurrentPage);
                    collectionRequest = pageList.NextPageRequest;
                }
                else
                {
                    break;
                }
            }
            return list.ToArray();
        }
        public static async Task<Group[]> GetAllAsync(this IGraphServiceGroupsCollectionRequest pagedCollectionRq)
        {
            var list = new List<Group>();

            var collectionRequest = pagedCollectionRq;
            while (true && collectionRequest != null)
            {
                var pageList = await collectionRequest.GetAsync();
                if (pageList.CurrentPage.Count > 0)
                {
                    list.AddRange(pageList.CurrentPage);
                    collectionRequest = pageList.NextPageRequest;
                }
                else
                {
                    break;
                }
            }
            return list.ToArray();
        }
        public static async Task<DirectoryObject[]> GetAllAsync(this IGroupMembersCollectionWithReferencesRequest pagedCollectionRq)
        {
            var list = new List<DirectoryObject>();

            var collectionRequest = pagedCollectionRq;
            while (true && collectionRequest != null)
            {
                var pageList = await collectionRequest.GetAsync();
                if (pageList.CurrentPage.Count > 0 )
                {
                    list.AddRange(pageList.CurrentPage);
                    collectionRequest = pageList.NextPageRequest;
                }
                else
                {
                    break;
                }
            }
            return list.ToArray();
        }
        public static async Task<SubscribedSku[]> GetAllAsync(this IGraphServiceSubscribedSkusCollectionRequest pagedCollectionRq)
        {
            var list = new List<SubscribedSku>();

            var collectionRequest = pagedCollectionRq;
            while (true && collectionRequest != null)
            {
                var pageList = await collectionRequest.GetAsync();
                if (pageList.CurrentPage.Count > 0)
                {
                    list.AddRange(pageList.CurrentPage);
                    collectionRequest = pageList.NextPageRequest;
                }
                else
                {
                    break;
                }
            }
            return list.ToArray();
        }

        public static async Task<Event[]> GetAllAsync(this IUserEventsCollectionRequest pagedCollectionRq)
        {
            var list = new List<Event>();
            var collectionRequest = pagedCollectionRq;
            while (true && collectionRequest != null)
            {
                var pageList = await collectionRequest.GetAsync();
                if (pageList.CurrentPage.Count > 0)
                {
                    list.AddRange(pageList.CurrentPage);
                    collectionRequest = pageList.NextPageRequest;
                }
                else
                {
                    break;
                }
            }
            return list.ToArray();
        }
        
    }
}