// =========================================================================
//  THIS CODE IS DEVELOPED BY KISHORE DHANEKULA
//
//  YOU HAVE COMPLETE RIGHTS TO MODIFY THIS FILE AS NEEDED FOR YOUR USAGE.
//  HOWEVER, IF YOU ARE DISTRIBUTING THIS FILE "AS IS", PROVIDE CREDITS
//  TO THE ORIGINAL AUTHOR.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =========================================================================
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;

namespace KD_Clone_Entities
{
    public class CloneParentAndChildEntities : CodeActivity
    {
        [Input("Child Entities: Entity name: Split by '|' e.g. account|contact")]
        public InArgument<string> ChildEntities_SchemaName { get; set; }

        [Input("Child Entities Parent Lookup Attribute Schema Name: Split by '|' e.g. accountid|contactid")]
        public InArgument<string> ChildEntities_ParentLookupAttributeSchemaName { get; set; }

        [Input("Entity with AutoNumber or Skip Attributes Schema Name: Split by '|' e.g. account;accountnumber|contact;contactnumber")]
        public InArgument<string> Entity_AutoNumberORSkipAttributesSchemaName { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService userService = serviceFactory.CreateOrganizationService(context.UserId);
            IOrganizationService systemService = serviceFactory.CreateOrganizationService(null);

            try
            {
                //Skip Attributes
                List<string> skipAttributes = new List<string> { "statecode", "statuscode" };
                //Child Entities
                string childEntities = ChildEntities_SchemaName.Get<string>(executionContext);
                List<string> childEntitiesList = string.IsNullOrEmpty(childEntities) ? new List<string>() : childEntities.Split('|').ToList();
                //Child Record Parent Lookups
                string childentity_ParentLookups = ChildEntities_ParentLookupAttributeSchemaName.Get<string>(executionContext);
                List<string> childentity_ParentLookupsList = string.IsNullOrEmpty(childentity_ParentLookups) ? new List<string>() : childentity_ParentLookups.Split('|').ToList();

                string entityWithAutonumbers = Entity_AutoNumberORSkipAttributesSchemaName.Get<string>(executionContext);
                List<string> entityWithAutonumbersList = string.IsNullOrEmpty(entityWithAutonumbers) ? new List<string>() : entityWithAutonumbers.Split('|').ToList();

                tracer.Trace("List of Entities: " + childEntitiesList.Count);
                tracer.Trace("List of ParentLookups: " + childentity_ParentLookupsList.Count);
                tracer.Trace("List of AutoNumbers: " + entityWithAutonumbersList.Count);

                //Get the Entity Attributes to be Cloned.
                var entityToBeCloned = systemService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));
                var newParentClone = CloneEntity(systemService, tracer, entityToBeCloned, skipAttributes, entityWithAutonumbersList);
                var newParentCloneId = systemService.Create(newParentClone);
                tracer.Trace("Parent Entity Cloning Complete");
                if (childEntitiesList != null && childEntitiesList.Count > 0 && childentity_ParentLookupsList != null && childentity_ParentLookupsList.Count > 0)
                {
                    tracer.Trace("Child Entities exists");
                    for (var i = 0; i < childEntitiesList.Count; i++)
                    {
                        var childRecords = RetrieveChildRecords(systemService, tracer, childEntitiesList[i], childentity_ParentLookupsList[i], context.PrimaryEntityId);
                        foreach (Entity childRecord in childRecords)
                        {
                            var newChild = CloneEntity(systemService, tracer, childRecord, skipAttributes, entityWithAutonumbersList);
                            newChild[childentity_ParentLookupsList[i]] = new EntityReference(context.PrimaryEntityName, newParentCloneId);
                            tracer.Trace("Child Entity Cloning Done, Creating");
                            systemService.Create(newChild);
                            tracer.Trace("Child Entity Created");
                        }
                    }
                }

            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        private Entity CloneEntity(IOrganizationService systemService, ITracingService tracer, Entity source, List<string> skipAttributes, List<string> entityWithAutoNumbers)
        {
            tracer.Trace("Cloning the Entity Schema Name: " + source.LogicalName);
            Entity newEntity = new Entity(source.LogicalName);
            foreach (var attribute in source.Attributes)
            {
                if (entityWithAutoNumbers != null && entityWithAutoNumbers.Count > 0 && entityWithAutoNumbers.Contains(source.LogicalName + ";" + attribute.Key))
                    tracer.Trace("Autonumber Key Exists: " + attribute.Key);
                if (skipAttributes.Contains(attribute.Key) || (entityWithAutoNumbers != null && entityWithAutoNumbers.Count > 0 && entityWithAutoNumbers.Contains(source.LogicalName + ";" + attribute.Key)) || attribute.Key == source.LogicalName + "id")
                    continue;               

                newEntity.Attributes.Add(attribute.Key, attribute.Value);
            }
            tracer.Trace("Cloning the Entity Complete");
            return newEntity;
        }

        private List<Entity> RetrieveChildRecords(IOrganizationService systemService, ITracingService tracer, string sourceEntitySchemaName, string parentLookupSchemaName, Guid parentLookupID)
        {
            QueryByAttribute querybyattribute = new QueryByAttribute(sourceEntitySchemaName)
            {
                ColumnSet = new ColumnSet(true),
                Attributes = { parentLookupSchemaName },
                Values = { parentLookupID }
            };
            var entityCollection = systemService.RetrieveMultiple(querybyattribute);
            return entityCollection.Entities.ToList();
        }
    }
}
