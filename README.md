# InsertOoxmlTest
Simple project to test the InsertOoxml method in Word Online

## Problem:

Also detailed on [Stack Overflow](https://stackoverflow.com/questions/52092107/word-insertooxml-method-returning-error-unknown-for-all-input)

When the InsertOoxml method is called with any valid OOxml input it fails with error:
```
word-web-16.00.debug.js:11162 Uncaught (in promise) Error: unknown at new RuntimeError (word-web-16.00.debug.js:11162) at RequestContext.ClientRequestContext.processRequestExecutorResponseMessage (word-web-16.00.debug.js:13713) at word-web-16.00.debug.js:13620
```

To avoid cofusing the issue with quesitons about the validitity of the Ooxml used to call the method, this project includes the following code behind a button in the UI that gets the current selection and attempts to duplicate the content via Ooxml.

``` javascript
await Word.run(async (context) => {

   var sourceRange = context.document.getSelection();
   var contentToCopy = sourceRange.getOoxml();
   await context.sync();

   sourceRange.insertOoxml(contentToCopy.value, 'After');
   await context.sync();
});
```

Just highlight some text in the document and press the button to replicate the issue.
