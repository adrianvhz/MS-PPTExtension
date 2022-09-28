npm run start:desktop // not working correctly (current code optimized for web). [Reason](https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/powerpoint-add-ins?view=powerpoint-js-1.4#detect-the-presentations-active-view-and-handle-the-activeviewchanged-event) (Detect the presentation's active view and handle the ActiveViewChanged event SECTION) 


npm run start:web        // Edit package.json  "scripts"->"start:web"   :     link = (your online office ppt project shared url). Create a office ppt [here]("https://www.office.com/launch/powerpoint?ui=es-ES&rs=CL&auth=1"). Then copy the shared link.


npm run stop:desktop

npm run stop:web

//

Or

//

npm run stop
