# msgraph-helper
Simplifies usage of the MSGraph client in your SPFX solutions.

## Usage
Import the helper:

```typescript
import MSGraph from 'msgraph-helper';
```

In your main webpart or extension in the `onInit` function, add the following:

```typescript
 await MSGraph.Init(this.context.msGraphClientFactory);
 ```

 Anywhere in your soluton you can now do:

 ```typescript
import MSGraph from 'msgraph-helper';

let memberOf = await MSGraph.Get('/me/memberOf');
```


The helper supports `Get`, `Patch`, `Put` and `Delete`.