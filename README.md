## Welcome to ChatExcel agent

<h4>This is a chat-based Excel data manipulation tool.<br>
You can ask for chat to perform various data manipulation tasks on your Excel files.<br></h4>

### Usage

```
python main.py
```

### Example of chat iteractions
```
- "Calculate the sum of the salaries in cell F3."
- "Make cell F2 the average of the salaries."
- "Insert the minimum age in cell F4."
- "Replace the head of column F by " "."
- "Exclude columns F."
- "Replace 50000 by 80000 in cel D5"
- "Make cell G1 with count Chicago in column C"

```

### Setup

1. First create an .env file with your model credentials.
<br>To do that rename the .env-example to .env and replace the variables values by your credentials. For example:
```
mv .env-example .env
vi .env:

MODEL = "gpt-4o-mini" # deployment model name
API_VERSION = "2024-06-01" # for Azure OpenAI Service
API_KEY = "your_api_key"  # your API key
ENDPOINT = "https://your-account-name.cognitiveservices.azure.com"
```

Observation: You can parse your AzureOpenAI credentials or the OpenAI ones. Always use the same variables name for those you have.


2. Create and activate a virtual env:
```
python -m venv .venv
.venv/Scripts/activate  # for windows
source .venv/bin/activate  # for linux
```

3. Install dependences:
    
```
pip install -r requirements.txt
```