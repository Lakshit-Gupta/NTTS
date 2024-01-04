from newsapi import NewsApiClient
import win32com.client as wincl

# Initialize the News API client
newsapi = NewsApiClient(api_key='23cb16f32acd4543b9d70a0089c1e594')
#Initializing Speech API Engine
spk = wincl.Dispatch("SAPI.SpVoice")
vcs = spk.GetVoices()
SVSFlag = 11


# Function to retrieve and read news articles by category
def retrieve_and_read_news_by_category(category, selected_voice_index,no_of_news):
    try:
        top_headlines = newsapi.get_top_headlines(country='in', category=category, language='en',page_size=no_of_news)
        articles = top_headlines['articles']
        if not articles:
            print(f"No top headlines found for '{category}'")
        else:
            for article in articles:
                print(f"Title: {article['title']}")
                print(f"Description: {article['description']}")
                print("----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
                text_to_speech(article['title'], selected_voice_index)
                text_to_speech(article['description'], selected_voice_index)
    except Exception as e:
        print(f"Error fetching top headlines: {e}")
# Function to retrieve and read news articles by source/string
def retrieve_and_read_news_by_source(source, selected_voice_index,no_of_news):
    try:
        q = input("Enter the type of news you want to see: ")

        all_articles = newsapi.get_everything(q=q, sources=source, language='en', sort_by='relevancy',page_size=no_of_news)
        articles = all_articles['articles']
        if not articles:
            print(f"No articles found for {source}.")
        else:
            for article in articles:
                print(f"Title: {article['title']}")
                print(f"Description: {article['description']}")
                print(
                    "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
                text_to_speech(article['title'],selected_voice_index)
                text_to_speech(article['description'],selected_voice_index)
    except Exception as e:
        print(f"Error fetching articles: {e}")


# Function to convert text to speech
def text_to_speech(text, selected_voice_index):
    try:
        if selected_voice_index < len(vcs):
            spk.Voice = vcs.Item(selected_voice_index)
        else:
            print("Selected voice not available. Using the default voice.")

        spk.Speak(text)
    except Exception as e1:
        print(f"Error has occurred {e1}")


# Main function to run the combined project
def main():
    print("Welcome to the news app")

    # Prompt the user to choose the voice
    voice_choice = input("Enter 0 for male voice or 1 for female voice: ")

    if voice_choice == '0':
        selected_voice_index = 0
    elif voice_choice == '1':
        selected_voice_index = 1
    else:
        print("Invalid choice for voice. Using the default voice.")
        selected_voice_index = 1  # Default to female voice

    choice = input("Enter 1 to search by category: (More Precise) or 2 to search by source/String: (More Versatile): ")
    no_of_news = int(input("Enter no of News You want to see"))

    if choice == '1':
        categories = {
            '1': 'business',
            '2': 'entertainment',
            '3': 'general',
            '4': 'health',
            '5': 'science',
            '6': 'sports',
            '7': 'technology',
        }
        print("Choose the type of news:")
        for key, value in categories.items():
            print(f"{key}. {value}")

        category_choice = input("Enter the category number: ")
        if category_choice not in categories:
            print("Invalid choice")
            return
        category = categories[category_choice]
        retrieve_and_read_news_by_category(category, selected_voice_index,no_of_news)

    elif choice == '2':
        sources = {
            '1': 'the-times-of-india',
            '2': 'bbc-news',
            '3': 'cnn',
            '4': 'business-insider',
            '5': 'ign',
            '6': 'national-geographic',
            '7': 'the-wall-street-journal',
        }
        print("Choose the news source:")
        for key, value in sources.items():
            print(f"{key}. {value}")

        source_choice = input("Enter the source number: ")
        if source_choice not in sources:
            print("Invalid choice")
            return
        source = sources[source_choice]
        retrieve_and_read_news_by_source(source,selected_voice_index,no_of_news)
    else:
        print("Invalid choice")


if __name__ == "__main__":
    main()
