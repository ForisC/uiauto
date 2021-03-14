# The uiauto-patch module

Windows uiautomatio module based on [uiautomation](https://pypi.org/project/uiautomation)


## New Featrues

* Support new `searchProperties`: `Child`<br>
  to find a control which contain specified child control<br>

  example:<br>

    ```
    target = uiauto.WindowControl(Name="AAA",
             Child=uiauto.Control(SubName="BBB"))
    ``` 

* Auto refind control while getting pattern<br>

    example:<br>
    1. Open a window named AAA
    2. ```
       target = WindowControl(Name="AAA")
       ```
    3. Close the AAA window and reopen it again, so the original object was already not exist
    4. ```
       target.GetTransformPattern().Move(0,0)
       ```
    5. It will be successful to get TransformPattern and to move the window even it is not the original window