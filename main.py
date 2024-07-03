import streamlit as st
# from langchain.agents import create_pandas_dataframe_agent
from langchain_experimental.agents import create_pandas_dataframe_agent
from langchain_openai import ChatOpenAI, AzureChatOpenAI
from dotenv import load_dotenv
import pandas as pd

def main():
    load_dotenv()

    # Set page configuration
    st.set_page_config(page_title='Ask Your Excel', page_icon='📊', layout='wide')

    # Add a title and description
    st.title("Ask Your Excel 📊")
    st.markdown("""
        **Upload an Excel file and ask questions about its contents.**
        This app uses a powerful AI to understand and analyze your data.
    """)

    # File uploader for Excel
    st.sidebar.header("Upload Your Excel File")
    user_excel = st.sidebar.file_uploader('Choose an Excel file', type=["xlsx", "xls"])

    # Instructions and example questions
    st.sidebar.markdown("""
        ### How to use:
        1. Upload your Excel file using the uploader above.
        2. Enter your question in the input box below.
        3. Click 'Submit' to get the answer.
        
        ### Example Questions:
        - What is the total sales for the month of June?
        - How many unique products are there?
        - What is the average price of products?
    """)
    max_tokens = st.sidebar.slider("Max Tokens", min_value=500, max_value=3000, value=1500, step=100)
    if user_excel is not None:
        st.subheader("Ask a Question About Your Excel Data")
        user_question = st.text_input("What do you want to know from your Excel?")

        if user_question:
            # Display the user's question
            st.write(f'**You asked:** {user_question}')
            try:
            # Read the Excel file into a DataFrame
                df = pd.read_excel(user_excel, sheet_name=None)

                # Combine all sheets into one DataFrame
                combined_df = pd.concat(df.values(), ignore_index=True)

                # Initialize the AI model
                llm = AzureChatOpenAI(
                    azure_deployment="gpt-4",
                    openai_api_version="2023-05-15",
                    azure_endpoint="https://aibcp4.openai.azure.com/",
                    max_tokens=max_tokens,
                    api_key="a9b5778f059648b7863c397ff8f8248a",
                )

                # Create a Pandas DataFrame agent
                agent = create_pandas_dataframe_agent(llm=llm, df=combined_df, verbose=True,allow_dangerous_code=True)

                # Get the response from the agent
                with st.spinner('Analyzing your data...'):
                    response = agent.run(user_question)

                # Display the response
                st.success("Here is the answer:")
                st.write(response)
            except Exception as e:
                st.error(f"An error occurred while processing the file: {str(e)}")
        else:
            st.info("Please enter a question to get an answer.")

if __name__ == '__main__':
    main()
