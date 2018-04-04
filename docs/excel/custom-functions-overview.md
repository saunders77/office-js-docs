# Create custom functions in Excel (Preview)

Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in. Users can then access custom functions like any other native function in Excel (like =SUM()). This article explains how to create custom functions in Excel.

The following illustration shows you how custom functions work in the Excel UI.

<img src="../../images/custom-function.gif" width="579" height="383" />

Here’s the code for a sample custom function that adds 42 to a pair of numbers.

```js
function add42 (a, b) {
    return a + b + 42;
}
```

Custom functions are now available in preview. Follow these steps to try them:

1.  Install Office 2016 for Windows and join the [Office Insider](https://products.office.com/en-us/office-insider) program.
2.  Clone the *Excel-Custom-Functions* repo and follow the instructions in *README.md* to start the add-in in Excel.
3.  Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.

See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.

## Learn the basics


In the cloned sample repo, you’ll see the following files:

-   *customfunctions.js*, which contains the custom function code to add to Excel.
-   *customfunctions.json*, which contains the registration code to connect your custom function to Excel. Registration makes your custom functions appear in the list of available functions displayed when users type in cells.
-   *customfunctions.html*, which provides a &lt;Script&gt; reference to the JS file. This file does not display UI in Excel.
-   *manifest.xml*, which tells Excel the location of your HTML, JS, and JSON files needed to run custom functions.

### JavaScript file (*customfunctions.js*)

The following code in customfunctions.js declares the custom function `ADD42`.

```js
function ADD42 (a, b) {
    return a + b + 42;
}
```
The following code in customfunctions.json declares the metadata for the same function:

{
    call: ADD42,
    description: "Adds 42 to the sum of two numbers",
    helpUrl: "https://www.contoso.com/help.html",
    result: {
        resultType: Excel.CustomFunctionValueType.number,
        resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    parameters: [{
        name: "num 1",
        description: "The first number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    },
    {
        name: "num 2",
        description: "The second number",
        valueType: Excel.CustomFunctionValueType.number,
        valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
    }],
    options:{ sync:true }
}

```

You need the following parameters to register the function in Excel:

-   Function name: The "name" defines the function name (in this case ADD42 is the function name). The prefix (like "CONTOSO", which appears before the name) is defined in the manifest. The prefix and the function name are separated using a period: to use your custom function, combine the function's prefix (CONTOSO) with the function's name (ADD42) and enter `=CONTOSO.ADD42` into a cell. By convention, prefixes and function names should use upper case letters. The prefix is intended to be used as an identifier for your add-in.
-   `description`: The description appears in the autocomplete menu in Excel.
-   `helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.
-   `result`: Defines the type of information returned by the function to Excel.

    -   `resultType`: Your function can return a `"string"`, `"number"`, or `"boolean"`. For more information see &lt;&lt;LINK&gt;&gt;.

    -   resultDimensionality: Your function can return either a single (`"scalar"`) value or a `"matrix"` of values. When returning a matrix of values, your function returns an array, where each array element is another array that represents a row of values. For more information, see &lt;&lt;LINK&gt;&gt;. The following example returns a 3-row, 2-column matrix of values from a custom function.

```js
return [["first","row"],["second","row"],["third","row"]];
```

-   Your custom function may take arguments as input. The arguments passed to your custom function are specified in the *parameters* property. The order of the parameters in the definition must match the order in the JavaScript function. For each parameter, define these properties:

    -   `name`: The string displayed in Excel to represent the parameter.
    -   `description`: The string displayed for more information about the parameter.
    -   `valueType`: A `"number"` or `"string"`, similar to the resultType property described earlier.
    -   `valueDimensionality`: A `"scalar"` value or `"matrix"` of values, similar to the resultDimensionality property described previously. Matrix-type parameters allow the user to select ranges larger than a single cell.

-   `options`: enables special types of custom functions that are described in more detail later in this article.

To complete registration of all functions defined using `Excel.Script.customFunctions`, ensure you call `CustomFunctions.addAll()`.

After registration, custom functions are available in all workbooks (not only the one where the add-in ran initially) for a user. The functions are displayed in the autocomplete menu when the user starts typing it.

### Manifest file (*manifest.xml*)

The following example in manifest.xml allows Excel to locate the code for your functions.

```xml

<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">

        <Hosts>
			<Host xsi:type="Workbook">
				<AllFormFactors>
					<ExtensionPoint xsi:type="CustomFunctions">
						<Script>
							<SourceLocation resid="residjs" />
						</Script>
						<Page>
							<SourceLocation resid="residhtml"/>
						</Page>
						<Metadata>
							<SourceLocation resid="residjson" />
						</Metadata>
						<Namespace resid="residNS" />
					</ExtensionPoint>
				</AllFormFactors>
			</Host>
		</Hosts>
		<Resources>
			<bt:Urls>
				<bt:Url id="residjson" DefaultValue="http://127.0.0.1:8080/customfunctions.json" />
				<bt:Url id="residjs" DefaultValue="http://127.0.0.1:8080/customfunctions.js" />
				<bt:Url id="residhtml" DefaultValue="http://127.0.0.1:8080/customfunctions.html" />
			</bt:Urls>
			<bt:ShortStrings>
				<bt:String id="residNS" DefaultValue="CONTOSO" />
			</bt:ShortStrings>
		</Resources>

</VersionOverrides>

```

The previous code specifies:

-   A &lt;`Script`&gt; element, which is only used for synchronous functions during the developer preview.
-   A &lt;`Page`&gt; element, which links to the HTML page of your add-in. The HTML page includes a &lt;Script&gt; reference to the JavaScript file (*customfunctions.js*) that contains the custom function and registration code. The HTML page is a hidden page and is never displayed in the UI. It's used for asynchronous functions during the developer preview
-   A &lt;`Metadata`&gt; element pointing to the JSON file.

## Asynchronous functions and synchronous functions

If your custom function retrieves data from the web, you need to make an asynchronous call to fetch it. When calling external web services, your custom function must:

1.   Return a JavaScript Promise to Excel.
2.   Make the http request to call the external service.
3.   Resolve the promise using the `setResult` callback. `setResult` sends the value to Excel.

The following code shows an example of a custom function that retrieves the temperature of a thermometer.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult, setError){
        sendWebRequestExample(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

Designate functions as asynchronous by setting the option `"sync": false` in the metadata file. During the developer preview, asynchronous functions run in a separate browser process, whereas synchronous functions run in the Excel process, allowing them to run much faster and to run concurrently.

## Streamed functions

Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations. For example, the `incrementValue` custom function in the following code adds a number to the result every second, and Excel displays each new value automatically using the `setResult` callback. To see the registration code used with `incrementValue`, read the *customfunctions.js* file.

```js
function incrementValue(increment, caller){ 
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

For streamed functions, the final parameter, `caller`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function. It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell. In order for Excel to pass the `setResult` function in the `caller` object, you must declare support for streaming during your function registration by setting the parameter `stream` to `true` in the metadata file.

## Cancellation

You can cancel streamed functions and asynchronous functions. Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load. Excel cancels function calls in the following situations:
- The user edits or deletes a cell that references the function.
- One of the arguments (inputs) for the function changes. In this case, a new function call is triggered in addition to the cancelation.
- The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.

The following code shows the previous example with cancellation implemented. In the code, the `caller` object contains an `onCanceled` function which should be defined for each custom function.

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## Saving state

Custom functions can save data in global JavaScript variables. In subsequent calls, your custom function may use the values saved in these variables. Saved state is useful when users enter multiple instances of the same custom function, and they need to share data with each other. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.

The following code shows an implementation of the previous temperature-streaming function that saves state using the `savedTemperatures` variable. The code demonstrates the following concepts:

-   **Saving data.** `refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second. New temperatures are saved in the savedTemperatures variable.

-   **Using saved data.** `streamTemperature` updates the temperature values displayed in the Excel UI every second. Temperatures are read from `savedTemperature`, and then sent to the Excel UI using `setResult`. Users may call `streamTemperature` from several cells in the Excel UI. Each call to `streamTemperature` will read data from `savedTemperatures`.

> In this case, we register `streamTemperature` as the custom function in Excel.

```js
var savedTemperatures{};

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID);
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequestExample(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## Working with ranges of data

Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.

For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel. The following function takes the parameter `temperatures`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.

```js
function secondHighestTemp(temperatures){ 
     var highest = -273, secondHighest = -273;
     for(var i = 0; i < temperatures.length; i++){
         for(var j = 0; j < temperatures[i].length; j++){
             if(temperatures[i][j] <= highest){
                 secondHighest = highest;
                 highest = temperatures[i][j];
             }
             else if(temperatures[i][j] <= secondHighest){
                 secondHighest = temperatures[i][j];
             }
         }
     }
     return secondHighest;
 }
```

## Known issues

The following features aren't yet supported in the Developer Preview.

-   Help URLs and parameter descriptions are not yet used by Excel.

-   Custom functions are not available on Excel for mobile clients or Excel Online.

-   Currently, add-ins rely on a hidden browser process to run asynchronous custom functions. In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory. Additionally, the HTML page referenced by the &lt;Page&gt; element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly. To prepare for this change, ensure your custom functions do not use the webpage DOM.

## Changelog

- **Nov 7, 2017**: Shipped the custom functions preview and samples
- **Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later
- **Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)
