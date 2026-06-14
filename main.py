import os
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain.agents import create_agent
from tools import web_search, search_images, generate_pdf_report


SYSTEM_PROMPT = """You are a helpful AI search assistant. Your job is to answer the user's question by:

1. Deciding if you need to search the web for current or factual information
2. Using the web_search tool when needed to find accurate, up-to-date answers
3. Synthesizing search results into a clear, concise answer — cite your sources with URLs
4. If you already know the answer confidently (e.g. capital cities, basic math), answer directly without searching
5. When the user asks for a "report", "write-up", "document", or wants to save findings, use the generate_pdf_report tool to create a formatted PDF. Provide a clear title and well-structured content with sections, bullet points, and source URLs.
   IMPORTANT: After generating a PDF, always show the user the exact full file path that the tool returns so they know exactly where to find it.
6. When creating a PDF report, you SHOULD also find relevant images to make the report richer:
   - Use search_images to find pictures related to the report topic
   - Pass the image URLs (comma-separated) to generate_pdf_report via the 'images' parameter
   - Use [IMG] markers in the content body where you want images to appear
   - If you don't specify [IMG] markers, images will be placed at the end of the report

Always be helpful, accurate, and cite your sources when using search results."""


def main():
    load_dotenv()

    llm = ChatOpenAI(
        model="deepseek-chat",
        temperature=0,
        base_url="https://api.deepseek.com",
        api_key=os.getenv("DEEPSEEK_API_KEY"),
    )
    tools = [web_search, search_images, generate_pdf_report]

    agent = create_agent(
        model=llm,
        tools=tools,
        system_prompt=SYSTEM_PROMPT,
    )

    messages = []

    print("=" * 60)
    print("AI Search Agent — type 'quit' to exit, 'clear' to reset")
    print("=" * 60)

    while True:
        try:
            query = input("\nYou: ").strip()
            if not query:
                continue
            if query.lower() == "quit":
                print("Goodbye!")
                break
            if query.lower() == "clear":
                messages = []
                print("[Chat history cleared]")
                continue

            messages.append({"role": "user", "content": query})

            result = agent.invoke({"messages": messages})
            answer = result["messages"][-1].content
            print(f"\nAgent: {answer}")

            messages = result["messages"]

        except KeyboardInterrupt:
            print("\nGoodbye!")
            break
        except Exception as e:
            print(f"\nError: {e}")


if __name__ == "__main__":
    main()
