# Custom Function metadata

When you include [custom functions](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview.md) in an Excel Add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file). This article describes the format of the JSON file with examples.

## The functions array

The metadata is a JSON object that contains a single `functions` property whose value is an array of objects. Each of these objects represents one custom function. The following are it's properties:

|  Property  |  Data Type  |  Required?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No  |  A description of the function that will appear in the Excel UI. For example, "Converts a Celsius value to Fahrenheit". |
|  `helpUrl`  |  string  |   No  |  URL where your users can get help about the function. (It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"  |
|  `name`  |  string  |  Yes  |  This is the name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function. It should be the same as the function's name where it is defined in the JavaScript, except that it should be all upper-case to conform with Excel function style. For example, if the function is `convertCelsiusToFahrenheit`, the name should be "CONVERTCELSIUSTOFAHRENHEIT". |
|  `options`  |  object  |  No  |  Configure how Excel processes the function. See below for details. |
|  `parameters`  |  array  |  No  |  Metadata about the parameters to the function. See below for details. |
|  `result`  |  object  |  Yes  |  Metadata about the value returned by the function. See below for details. |

## The options object

The `options` object can configure how Excel processes the function. It has the following properties.

|  Property  |  Data Type  |  Required?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  No, default is `false`.  |  If `true`, Excel calls the `onCancelled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing the cell that references the function. To use this option, the last parameter of the function must be an object that represents the caller. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCancelled` member. Note, `cancelable` and `sync` may not both be `true`.  |
|  `stream`  |  boolean  |  No, default is `false`.  |  If `true`, the output is written repeatedly to the cell. Useful for rapidly changing data sources, such as a stock price ticker. To use this option, the last parameter of the function must be an object that represents the caller. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed to the `caller.setResult` method. Note, `stream` and `sync` may not both be `true`.|
|  `sync`  |  boolean  |  No, default is `false`  |  If `true`, the function runs synchronously. Note, `sync`  may not be `true` if either `cancelable` or `stream` are `true`.  |
|  `volatile`  |  boolean  |  No, default is `false`.  |  If `true`, the function re-executes each time calculation runs in the workbook. |

## The parameters array

The `parameters` property is an array of objects. Each of these objects represents a parameter. The objects have the following properties.

|  Property  |  Data Type  |  Required?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No |  A description of the parameter.  |
|  `dimensionality`  |  string  |  Yes  |  Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of any number of dimensions.  |
|  `name`  |  string  |  Yes  |  The name of the parameter.  |
|  `optional`  |  boolean  |  No, default is `false`.  |  If `true`, the parameter is not required.  |
|  `type`  |  string  |  Yes  |  The data type of the parameter. Must be "boolean", "isodate", "number", or "string".  |

## The result object

The `results` property provides metadata about the value returned from the function. It has the following properties.

|  Property  |  Data Type  |  Required?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  No  |  Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of any number of dimensions.  |
|  `type`  |  string  |  Yes  |  The data type of the parameter. Must be "boolean", "isodate", "number", or "string".  |

## Example

The following is an example of a metadata file for custom functions.

```json
{
	"functions": [
		{
			"name": "ADD42", 
			"description":  "Adds 42 to the input number",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "number",
				"dimensionality": "scalar"
			},
			"parameters": [
				{
					"name": "num",
					"description": "Number",
					"type": "number",
					"dimensionality": "scalar"
				}
			],
			"options": {
				"sync": true
			}
		},
		{
			"name": "ADD42WAIT", 
			"description":  "asynchronously wait 250ms, then add 42",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "number",
				"dimensionality": "scalar"
			},
			"parameters": [
				{
					"name": "num",
					"description": "Number",
					"type": "number",
					"dimensionality": "scalar"
				}
			],
			"options": {
				"sync": false
			}
		},
		{
			"name": "ISEVEN", 
			"description":  "Determines whether a number is even",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "boolean",
				"dimensionality": "scalar"
			},
			"parameters": [
				{
					"name": "num",
					"description": "the number to be evaluated",
					"type": "number",
					"dimensionality": "scalar"
				}
			],
			"options": {
				"sync": true
			}
		},
		{
			"name": "GETDAY",
			"description": "Gets the day of the week",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "string"
			},
			"parameters": [],
			"options": {
				"sync": true
			}
		},
		{
			"name": "INCREMENTVALUE", 
			"description":  "Counts up from zero",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "number",
				"dimensionality": "scalar"
			},
			"parameters": [
				{
					"name": "increment",
					"description": "the number to be added each time",
					"type": "number",
					"dimensionality": "scalar"
				}
			],
			"options": {
				"sync": false,
				"stream": true,
				"cancelable": true
			}
		},
		{
			"name": "SECONDHIGHEST", 
			"description":  "gets the second highest number from a range",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "number",
				"dimensionality": "scalar"
			},
			"parameters": [
				{
					"name": "range",
					"description": "the input range",
					"type": "number",
					"dimensionality": "matrix"
				}
			],
			"options": {
				"sync": true
			}
		}
	]
}

```
