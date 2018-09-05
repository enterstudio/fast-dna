import "jest";
import {
    getBreadcrumbs,
    getNavigation,
    HandleBreadcrumbClick,
    IBreadcrumbItem,
    INavigationItem,
    mapSchemaLocationFromDataLocation
} from "./form.utilities";
import {
    BreadcrumbItemEventHandler
} from "./form.props";
import * as alignHorizontalSchema from "../../app/components/align-horizontal/align-horizontal.schema.json";
import * as arraysSchema from "../../app/components/arrays/arrays.schema.json";
import * as objectsSchema from "../../app/components/objects/objects.schema.json";
import * as anyOfSchema from "../../app/components/any-of/any-of.schema.json";
import * as childrenSchema from "../../app/components/children/children.schema.json";
import * as textFieldSchema from "../../app/components/text-field/text-field.schema.json";

/**
 * Map schema location from data location
 */
describe("Map schema location from data location", () => {
    test("should return a schema location from a root data location", () => {
        const schemaLocation: string = mapSchemaLocationFromDataLocation("", {}, alignHorizontalSchema);

        expect(schemaLocation).toBe("");
    });
    test("should return a schema location from a nested property", () => {
        const schemaLocation: string = mapSchemaLocationFromDataLocation(
            "alignHorizontal",
            {alignHorizontal: "left"},
            alignHorizontalSchema
        );

        expect(schemaLocation).toBe("properties.alignHorizontal");
    });
    test("should return a schema location from an array", () => {
        const schemaLocation: string = mapSchemaLocationFromDataLocation("strings[0]", {strings: ["a"]}, arraysSchema);

        expect(schemaLocation).toBe("properties.strings");
    });
    test("should return a schema location from a nested array item", () => {
        const schemaLocation: string = mapSchemaLocationFromDataLocation(
            "objects[1].string",
            {objects: [{ string: "foo" }, { string: "bar" }]},
            arraysSchema
        );

        expect(schemaLocation).toBe("properties.objects.items.properties.string");
    });
    test("should return a schema location from anyOf/oneOf locations", () => {
        const schemaLocationRoot: string = mapSchemaLocationFromDataLocation("", {number: 5}, anyOfSchema);
        const schemaLocation: string = mapSchemaLocationFromDataLocation("number", {number: 5}, anyOfSchema);

        expect(schemaLocationRoot).toBe("");
        expect(schemaLocation).toBe("anyOf.1.properties.number");
    });
    test("should return a schema location from a nested anyOf/oneOf location", () => {
        const schemaLocationRootProperty: string = mapSchemaLocationFromDataLocation(
            "nestedAnyOf",
            {nestedAnyOf: {string: "foo"}},
            anyOfSchema
        );
        const schemaLocation: string = mapSchemaLocationFromDataLocation(
            "nestedAnyOf.string",
            {nestedAnyOf: {string: "foo"}},
            anyOfSchema
        );

        expect(schemaLocationRootProperty).toBe("anyOf.2.properties.nestedAnyOf");
        expect(schemaLocation).toBe("anyOf.2.properties.nestedAnyOf.anyOf.1.properties.string");
    });
});

/**
 * Gets the navigation
 */
describe("getNavigation", () => {
    test("should return a single navigation item when the location is at the root", () => {
        const navigation: INavigationItem[] = getNavigation(
            "",
            {
                alignHorizontal: "left"
            },
            alignHorizontalSchema
        );

        expect(navigation.length).toBe(1);
        expect(navigation[0].dataLocation).toBe("");
        expect(navigation[0].schema).toEqual(alignHorizontalSchema);
        expect(navigation[0].data).toEqual({
            alignHorizontal: "left"
        });
    });
    test("should return navigation items for a nested property", () => {
        const navigation: INavigationItem[] = getNavigation(
            "optionalObjectWithNestedObject.nestedObject",
            {
                optionalObjectWithNestedObject: {
                    nestedObject: {
                        boolean: true
                    }
                }
            },
            objectsSchema
        );

        expect(navigation.length).toBe(3);
        expect(navigation[0].dataLocation).toBe("");
        expect(navigation[0].schema).toEqual(objectsSchema);
        expect(navigation[0].data).toEqual({
            optionalObjectWithNestedObject: {
                nestedObject: {
                    boolean: true
                }
            }
        });
        expect(navigation[1].dataLocation).toBe("optionalObjectWithNestedObject");
        expect(navigation[1].schema).toEqual(objectsSchema.properties.optionalObjectWithNestedObject);
        expect(navigation[1].data).toEqual({
            nestedObject: {
                boolean: true
            }
        });
        expect(navigation[2].dataLocation).toBe("optionalObjectWithNestedObject.nestedObject");
        expect(navigation[2].schema).toEqual(objectsSchema.properties.optionalObjectWithNestedObject.properties.nestedObject);
        expect(navigation[2].data).toEqual({
            boolean: true
        });
    });
    test("should return navigation items for an array", () => {
        const navigation: INavigationItem[] = getNavigation(
            "objects.1",
            {objects: [{ string: "foo" }, { string: "bar" }]},
            arraysSchema
        );

        expect(navigation.length).toBe(2);
        expect(navigation[0].dataLocation).toBe("");
        expect(navigation[0].schema).toEqual(arraysSchema);
        expect(navigation[0].data).toEqual({objects: [{ string: "foo" }, { string: "bar" }]});
        expect(navigation[1].dataLocation).toBe("objects[1]");
        expect(navigation[1].schema).toEqual(arraysSchema.properties.objects);
        expect(navigation[1].data).toEqual({ string: "bar" });
    });
    test("should return navigation items for a anyOf/oneOfs", () => {
        const navigationRoot: INavigationItem[] = getNavigation(
            "",
            {nestedAnyOf: {string: "foo"}},
            anyOfSchema
        );
        const navigation: INavigationItem[] = getNavigation(
            "nestedAnyOf",
            {nestedAnyOf: {string: "foo"}},
            anyOfSchema
        );

        expect(navigationRoot.length).toBe(1);
        expect(navigation[0].dataLocation).toBe("");
        expect(navigation[0].schema).toEqual(anyOfSchema);
        expect(navigation[0].data).toEqual({nestedAnyOf: {string: "foo"}});

        expect(navigation.length).toBe(2);
        expect(navigation[0].dataLocation).toBe("");
        expect(navigation[0].schema).toEqual(anyOfSchema);
        expect(navigation[0].data).toEqual({nestedAnyOf: {string: "foo"}});
        expect(navigation[1].dataLocation).toBe("nestedAnyOf");
        expect(navigation[1].schema).toEqual(anyOfSchema.anyOf[2].properties.nestedAnyOf);
        expect(navigation[1].data).toEqual({string: "foo"});
    });
});

/**
 * Gets breadcrumbs from navigation items
 */
describe("Get the breadcrumbs", () => {
    const handleBreadcrumbClick: HandleBreadcrumbClick = (
        schemaLocation: string,
        dataLocation: string,
        schema: any
    ): BreadcrumbItemEventHandler => {
        return (e: React.MouseEvent): void => {
            e.preventDefault();
        };
    };

    test("should return a single breadcrumb item", () => {
        const navigation: INavigationItem[] = getNavigation(
            "",
            {
                alignHorizontal: "left"
            },
            alignHorizontalSchema
        );
        const breadcrumbs: IBreadcrumbItem[] = getBreadcrumbs(navigation, handleBreadcrumbClick);

        expect(breadcrumbs.length).toBe(1);
        expect(breadcrumbs[0].href).toBe("");
        expect(breadcrumbs[0].text).toBe("Component with align horizontal");
    });
    test("should return breadcrumbs for nested property locations", () => {
        const navigation: INavigationItem[] = getNavigation(
            "optionalObjectWithNestedObject.nestedObject",
            {
                optionalObjectWithNestedObject: {
                    nestedObject: {
                        boolean: true
                    }
                }
            },
            objectsSchema
        );
        const breadcrumbs: IBreadcrumbItem[] = getBreadcrumbs(navigation, handleBreadcrumbClick);

        expect(breadcrumbs.length).toBe(3);
        expect(breadcrumbs[0].href).toBe("");
        expect(breadcrumbs[0].text).toBe("Component with objects");
        expect(breadcrumbs[1].href).toBe("optionalObjectWithNestedObject");
        expect(breadcrumbs[1].text).toBe("object with nested object");
        expect(breadcrumbs[2].href).toBe("optionalObjectWithNestedObject.nestedObject");
        expect(breadcrumbs[2].text).toBe("Nested object");
    });
    test("should return breadcrumb items for an array location", () => {
        const navigation: INavigationItem[] = getNavigation(
            "objects.1",
            {objects: [{ string: "foo" }, { string: "bar" }]},
            arraysSchema
        );
        const breadcrumbs: IBreadcrumbItem[] = getBreadcrumbs(navigation, handleBreadcrumbClick);

        expect(breadcrumbs.length).toBe(2);
        expect(breadcrumbs[0].href).toBe("");
        expect(breadcrumbs[0].text).toBe("Component with array");
        expect(breadcrumbs[1].href).toBe("objects[1]");
        expect(breadcrumbs[1].text).toBe("Array");
    });
    test("should return items for an anyOf/oneOf location", () => {
        const navigationRoot: INavigationItem[] = getNavigation(
            "",
            {nestedAnyOf: {string: "foo"}},
            anyOfSchema
        );
        const navigation: INavigationItem[] = getNavigation(
            "nestedAnyOf",
            {nestedAnyOf: {string: "foo"}},
            anyOfSchema
        );
        const breadcrumbsRoot: IBreadcrumbItem[] = getBreadcrumbs(navigationRoot, handleBreadcrumbClick);
        const breadcrumbs: IBreadcrumbItem[] = getBreadcrumbs(navigation, handleBreadcrumbClick);

        expect(breadcrumbsRoot.length).toBe(1);
        expect(breadcrumbs[0].href).toBe("");
        expect(breadcrumbs[0].text).toBe("Component with anyOf");

        expect(breadcrumbs.length).toBe(2);
        expect(breadcrumbs[0].href).toBe("");
        expect(breadcrumbs[0].text).toBe("Component with anyOf");
        expect(breadcrumbs[1].href).toBe("nestedAnyOf");
        expect(breadcrumbs[1].text).toBe("Nested anyOf");
    });
});
