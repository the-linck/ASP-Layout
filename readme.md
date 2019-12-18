# ASP Layout

A minimalist Layout engine to use on Classic ASP projects, inspired by ASP.Net MVC projects but (obviously) a lot simpler.

This library uses ASP [ASP-Dynamic-Include]() to call Templates and Views set on the provided configuration variables.

As the mentionend base-project, there are not many relevent new resources to be added here, so certainly it won't be updated frequenly - except in case of errors or problems.


## How to use

You can either include _ASP-Dynamic-Include/DynamicInclude.asp_ and _Layout.asp_ manually on your project or just include the convenience-file _Bundle.asp_ to automatically have both libs at once.

Then, like in a MVC View, just set the Template to use (if you want), do your logic and call _View()_ method when you want to show your page's content.

## Provided methods and Collections

This methods come with the lib:

* **AddLibraries**(string|string[] **Scripts**)  
Adds scripts to be loaded as a libraries (on &lt;head&gt;)
* **AddPlugins**(string|string[] **Scripts**)  
Adds scripts to be loaded as a plugins (on &lt;body&gt;, after page's content).
* **AfterView**()  
Clears variables after the View is executed, preparing the script to end.
* **BeforeView**()  
Anything to be executed right before the View is executed (nothing by default).
* **GetScripts**(string|string[]|Dictionary **Files**)  
Prints all given CSS/JS scripts in HTML tags.
* string **ScriptName**(string **File**)  
Gets the file name of the File on provided path.
* **View**()  
Displays the configured View file on current Template.  
*If no Template is set, show the default blank HTML template.*  
*If the configured View doesn't exist, shows only the template.*

There are also the following Collections:

* **AddLibraries**(string|string[] **Scripts**)  
Adds scripts to be loaded as a libraries (on &lt;head&gt;)
* **AddPlugins**(string|string[] **Scripts**)  
Adds scripts to be loaded as a plugins (on &lt;body&gt;, after page's content).
* **AfterView**()  
Clears variables after the View is executed, preparing the script to end.
* **ViewData**()  
Settings of the layout engine.
    * **ViewData(*"libraries"*)**()  
    Scripts to be loaded on &lt;head&gt;.
    * **ViewData(*"plugins"*)**()  
    Scripts to be loaded on &lt;body&gt;, after page's content.
* **ViewBag**  
Arbitrary user data.


## Config and Info

All data tha is used to control this Layout engine behavior are keys of ViewData collection:

* string **ViewData(*"charset"*)**  
Defines the HTML header for page charset. Must be manually set.
* bool **ViewData(*"no-cache"*)**  
If the page and scripts must avoid cache.  
When true, ASP headers will be set to avoid cache and scripts will be loaded with a timestamp querystring.
* string **ViewData(*"page"*)**  
Current page filename (without extension).  
It's used to search the view.
* string **ViewData(*"path"*)**  
Relative path of the current script.
* string **ViewData(*"ready"*)**  
Exact time when the library finished loading, in database timestamp format.
* Date **ViewData(*"ready-timestamp"*)**  
Exact time when the library finished loading.
* string **ViewData(*"script"*)**  
Current page filename (with extension).
* string **ViewData(*"template"*)**  
Template file used to generate the page.  
When it's an empty string (default), the default blank HTML template is used.
* string **ViewData(*"title"*)**  
Page title, used for HTML &lt;head&gt;.



## Extending

You can easily extend this lib without even looking at it's code all, just going by one of this four ways - from the easiest to the most flexible:

* **Changing *ViewData*'s values**  
Just add new keys with custom settings to your templates.
* **Overriding the provided functions**  
Any function in VBScript can be overriden just declaring a new one with the same name after.
* **Including the lib in a custom Layout script**  
To keep a great performance it's better to use a ASP Include instead of Dynamic Include's functions
* **All of the previous**