import os
import warnings
import pandas as pd
from dotenv import load_dotenv, find_dotenv
from langchain_openai import AzureChatOpenAI, ChatOpenAI
from langchain_experimental.agents.agent_toolkits import create_pandas_dataframe_agent

warnings.filterwarnings("ignore", category=FutureWarning)

## Load environment variables from .env filetry:
try:
    load_dotenv(find_dotenv())
except Exception as e:
    print(f"Error loading .env file: {e}")
    raise

# Initialize the OpenAI model with the API key from the environment variable
MODEL = os.getenv('MODEL')
API_VERSION = os.getenv('API_VERSION')
API_KEY = os.getenv('API_KEY')
ENDPOINT = os.getenv('ENDPOINT')

try:
    try:    
        model = AzureChatOpenAI(model=MODEL, api_version=API_VERSION, 
                                api_key=API_KEY, azure_endpoint=ENDPOINT,
                                temperature=0)
    except ImportError:
        print("AzureChatOpenAI is not available. Falling back to ChatOpenAI.")
        # Fallback to ChatOpenAI if AzureChatOpenAI is not available
        model = ChatOpenAI(model=MODEL, api_key=API_KEY, temperature=0)
except:
    # Check if the required environment variables are set
    if not all([MODEL, API_KEY]):
        raise ValueError("Missing required environment variables.")

# Prompt for the agent to follow
# This prompt is used to instruct the agent on how to handle the data manipulation tasks.
prompt = """You are a data manipulator.
To locate correctly the row and column to update or insert the calculated value, following the rules: 
    - The first row is the header.
    - The pair letter-number represent the column and row index respectively.
    - Columns are in alphabetic order, so you should convert the letter column index to numerical \
        considering index 0 for column A, index 1 for column B, index 3 for column C, and so on.
    - Rows are numericals but you should convert the number to correct rows index as 0 for row 1, \
        1 for row 2, 2 for row 3, and so on.
Following the example to insert the calculated values into the correct position:
- Query: Make cell F2 the average of salaries
- Thought: 
    - First identify the column representing salaries. 
    - Now identify the column and row index assuming "F" equal column index 5 and "2" equal row index 1.
    - So now insert the average of salaries in iloc[1,5].
- You should do: 
    c = 5
    r = 1
    df = pd.read_excel({filename})
    salary_average = df['Salary'].mean()
    df.iloc[r,c] = salary_average
    df = df.rename(columns={df.columns[c]: ''})
    df.to_excel({filename}, index=False)
Warnings:
1. If the specified column or row (F2 equal iloc[1,5]) does not exist, the agent should create new ones to solve the query.
2. Never insert column name if is not required by the user query.
3. Make a double check on the actions before finish.
- 
"""

def load_file_and_agent(file: str, prompt: str) -> pd.DataFrame:
    df = pd.read_excel(file)
    print(f"Loaded {file}")
    print(df.head())  # Display the first few rows of the DataFrame
    prompt_ = prompt.replace("{filename}", file)

    # Create the agent with extra tools
    return create_pandas_dataframe_agent(
        model,
        df,
        agent_type="tool-calling",
        prefix=prompt_,
        suffix="Provide the final answer in a clear and structured format.",
        #extra_tools=extra_tools,  # Add custom tools
        allow_dangerous_code=True,
        verbose=False,
        max_iterations=30,
        max_execution_time=240
    )

def chat_excel(prompt: str) -> None:
    """Main function to run the ChatExcel application."""
    ### Load the Excel file
    print("Welcome to ChatExcel!")
    #print("This is a chat-based Excel data manipulation tool.")
    print("You can ask me to perform various data manipulation tasks on your Excel files.")
    print("Type 'exit' to quit the application.")
    file = input("Please enter the path to your Excel file: ")

    while not os.path.isfile(file):
        if file.lower() == "exit":
            return "Exiting the application."
        print(f"File {file} does not exist.")
        file = input("Please enter with the valid filepath: ")     

    agent = load_file_and_agent(file, prompt)
    print("Agent was created. You can now ask me to perform data manipulation tasks.\n")
    while True:
        user_input = input("Query: ")
        if user_input.lower() == "exit":
            print("Exiting the application.\nBye!")
            break

        response = agent.invoke(user_input)
        print(response["output"])

        agent = load_file_and_agent(file, prompt)


if __name__ == "__main__":
    chat_excel(prompt)