import time
from selenium import webdriver
import json
import pandas as pd
import streamlit as st


def test_args():
    options = webdriver.EdgeOptions()

    # This keeps the browser open after the script finishes
    options.add_experimental_option("detach", True)

    options.add_argument("--start-maximized")

    driver = webdriver.Edge(options=options)

    #launches the browser to navigate to the specified URL. Does not return anything.
    driver.get('https://app.fieldnation.com/workorders?list=workorders_available')

    #print(browser_content)

    # Keeping the browser open for demonstration purposes
    # input("Press Enter to close the browser...")

    # is intended to close the browser. If you have this command at the end of your script, it will close the browser regardless of the detach option
    # driver.quit()






# Open the browser and the profile using terminal. 
# navigate to C:\Program Files (x86)\Microsoft\Edge\Application
# msedge.exe  --user-data-dir="C:\Users\vertebrae\AppData\Local\Microsoft\Edge\User Data\Afriqo_Automation"
# Open website you want to automate and login manually to activate a user session. Now CLOSE this browser and run your script.
# The script will open the browser automaticall and run your automation. 
def open_existing_session():

    # Specify the user data directory to reuse the session
    options = webdriver.EdgeOptions()

    # This keeps the browser open after the script finishes
    options.add_experimental_option("detach", True)

    options.add_argument("--start-maximized")

    options.add_argument("user-data-dir=C:/Users/vertebrae/AppData/Local/Microsoft/Edge/User Data/Afriqo_Automation")
    # options.add_argument("profile-directory=Afriqo_Automation")  # Use "Default" or your specific profile

    driver = webdriver.Edge(options=options)

    # Adding a delay to ensure the browser starts correctly
    # time.sleep(5)
    
    #launches the browser to navigate to the specified URL. Does not return anything
    driver.get('https://app.fieldnation.com/workorders/')

    page_source = driver.page_source
    print(page_source)  # Prints the HTML source of the page


    # Your automation code here

    # driver.quit()



# For this function open the browser and the profile using terminal. Make sure to include the port.
# navigate to C:\Program Files (x86)\Microsoft\Edge\Application
# msedge.exe --remote-debugging-port=9444 --user-data-dir="C:\Users\vertebrae\AppData\Local\Microsoft\Edge\User Data\Afriqo_Automation"
# and login to your intended site. Do not close the browser. Now run your script
def open_existing_session_v2():

    # Specify the user data directory to reuse the session
    options = webdriver.EdgeOptions()

    options.add_experimental_option("debuggerAddress", "localhost:9444")

    driver = webdriver.Edge(options=options)

    # Adding a delay to ensure the browser starts correctly
    # time.sleep(5)
    
    #launches the browser to navigate to the specified URL. Does not return anything.
    driver.get('https://app.fieldnation.com/workorders/')


    # Your automation code here
    ######################################### Print page source
    # page_source = driver.page_source
    # # print(page_source)  # Prints the HTML source of the page

    # # Specify the path and name of the file where you want to save the HTML
    # file_path = 'page_source.html'

    # # Open the file in write mode and save the page source
    # with open(file_path, 'w', encoding='utf-8') as file:
    #     file.write(page_source)

    # # Print confirmation
    # print(f"Page source has been written to {file_path}")

    ############################################## Make ajax request

    # Define the JavaScript code for the AJAX request
    # ajax_request_js = """
    # var xhr = new XMLHttpRequest();
    # xhr.open('GET', 'https://app.fieldnation.com/v2/users/748312/tax', false);  // Use 'false' for synchronous request
    # xhr.send(null);
    # return xhr.responseText;
    # """

    # # Execute the JavaScript to make the AJAX request
    # response = driver.execute_script(ajax_request_js)

    # # Parse the JSON response (if applicable)
    # response_data = json.loads(response)

    # # Print the response
    # print(json.dumps(response_data, indent=4))

    ############################################# console.log a window.work_orders

    # Use execute_script to retrieve the work_orders object from the window
    work_orders = driver.execute_script("return window.work_orders;")

    # Optionally, pretty-print the JSON data if it's in JSON format
    #print(json.dumps(work_orders, indent=4))

    # Convert the work_orders object to a JSON string (if it is a dictionary or list)
    # By setting indent=4, the output JSON string will be pretty-printed with each nested level indented by 4 spaces, making it more human-readable.
    work_orders_json = json.dumps(work_orders, indent=4)

    # Specify the path and name of the file where you want to save the HTML
    work_order_file_path = 'work_orders_output.json'

    # Open the file in write mode and save the page source
    with open(work_order_file_path, 'w', encoding='utf-8') as file:
        file.write(work_orders_json)

    # Print confirmation
    print(f"Page source has been written to {work_order_file_path}")


    # Parse the JSON-like data into a Python dictionary
    data = json.loads(work_orders_json)

    # Normalize the JSON data into a flat table
    df = pd.json_normalize(data["results"])

    # Show the first few rows of the DataFrame
    # print(df.head())


    ###when you have a cell that has a complex structure
    # # 1 If you prefer to keep the nested structures but want to display them in a readable format, you can convert these objects to strings.
    # df['company.features'] = df['company.features'].apply(lambda x: json.dumps(x, indent=4))

    # # 2 Expand Lists into Rows: Example: Expanding a list column 'actions' into multiple rows
    # df = df.explode('company.features')

    ###


    ### Stremlit
    # Display the DataFrame using Streamlit
    st.write("Here is the DataFrame:")
    st.dataframe(df)  # You can also use st.table(df) for a static table display

    ###


    ##############################################
    # driver.quit()


open_existing_session_v2()
# test_args()

