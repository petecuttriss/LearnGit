export class Utils {
    
    public static groupBy<T extends any, K extends keyof T>(array: T[], key: K | { (obj: T): string }): Record<string, T[]> {
        const keyFn = key instanceof Function ? key : (obj: T) => obj[key];
        return array.reduce(
            (objectsByKeyValue, obj) => {
                const value = keyFn(obj);
                objectsByKeyValue[value] = (objectsByKeyValue[value] || []).concat(obj);
                return objectsByKeyValue;
            },
            {} as Record<string, T[]>
        );
    }

    public static sortByKey(unsorted: Object): Object {
        let sortedObject = {};
        Object.keys(unsorted).sort().forEach((key) => {
            sortedObject[key] = unsorted[key];
        });
        return sortedObject;
    }
}