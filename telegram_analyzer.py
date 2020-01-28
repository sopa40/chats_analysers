import json
import sys
from os import path
from collections import Counter

'''     
    structure of telegram JSON:
    about
    personal_information
    profile_pictures
    contacts
    frequent_contacts
    chats
'''

# for opening file
file_path = path.relpath("result.json")

# for saving all info about chats
contacts = {}

# account owner's id
owner_id = None

chats_to_analyze = 10

most_popular_words_number = 10

most_popular_symb_number = 10

last_sender = None

last_time_sent = None


def get_owner_name(data):
    personal_info = data.get("personal_information")
    if personal_info is None:
        print("Something went wrong while obtaining personal information. Exit")
        sys.exit()
    owner_first_name = personal_info.get("first_name")
    if owner_first_name is None:
        print("Something went wrong while obtaining owner's first name. Exit")
        sys.exit()
    owner_last_name = personal_info.get("last_name")
    if owner_last_name is None:
        print("Something went wrong while obtaining owner's last name. Exit")
        sys.exit()
    return owner_first_name + " "  + owner_last_name


def get_owner_id(data):
    personal_info = data.get("personal_information")
    if (personal_info is None):
        print("Something went wrong while obtaining personal information. Exit")
        sys.exit()
    owner_id = personal_info.get("user_id")
    if (owner_id is None):
        print("Something went wrong while obtaining owner's id. Exit")
        sys.exit()
    return owner_id


# sorting by number of messages
def insertionSort(arr):
    for i in range(1, len(arr)):
        temp = arr[i]
        key = len(get_messages(arr[i]))
        j = i - 1
        while j >= 0 and key > len(get_messages(arr[j])):
            arr[j + 1] = arr[j]
            j -= 1 
        arr[j + 1] = temp
    return arr


def get_chats(data):
    chats = data.get("chats")
    if chats is None:
        print("Something went wrong while extracting chats. Exit...")
        sys.exit()
    result = chats.get("list")
    if result is None:
        print("Something went wrong while getting list of chats. Exit...")
        sys.exit()
    return result


def get_chat_type(chat_info):
    result = chat_info.get("type")
    if result is None:
        print("Something went wrong while detecting chat type. Exit...")
        sys.exit()
    return result


def get_name(chat_info):
    user_name = chat_info.get("name")
    if user_name is None:
        print("Something went wrong while getting name. Id is", get_id(chat_info), ". Exit...")
        sys.exit()
    return user_name


def get_id(chat_info):
    user_id = chat_info.get("id")
    if user_id is None:
        print("Something went wrong while getting id. Exit...")
        sys.exit()
    return user_id


def get_name_by_id(id):
    if (id in contacts):
        return contacts[id]
    print("No such id in contacts! Exit...")
    sys.exit()


def get_messages(chat_info):
    result = chat_info.get("messages")
    if result is None:
        name = get_name(chat_info)
        print("Something went wrong while getting messages of ", name)
        sys.exit()
    return result


def init_owner_contact(id, name):
    global contacts
    contacts[id] = {}
    contacts[id]["name"] = name


def init_contact(id, name):
    global contacts
    contacts[id] = {}
    contact = contacts[id]
    contact["name"] = name
    contact["message_counter"] = 0
    contact["symbol_counter"] = 0
    contact["word_counter"] = 0
    contact["avg_len"] = 0
    contact['comma_ctr'] = 0
    contact['links_shared'] = 0
    contact['msg_edited'] = 0
    contact["words_usage"] = Counter()
    contact["symb_usage"] = Counter()
    contacts[owner_id][id] = {}
    contacts[owner_id][id]["message_counter"] = 0
    contacts[owner_id][id]["symbol_counter"] = 0
    contacts[owner_id][id]["word_counter"] = 0
    contacts[owner_id][id]["avg_len"] = 0
    contacts[owner_id][id]["comma_ctr"] = 0
    contacts[owner_id][id]["links_shared"] = 0
    contacts[owner_id][id]["msg_edited"] = 0
    contacts[owner_id][id]["words_usage"] = Counter()
    contacts[owner_id][id]["symb_usage"] = Counter()


def register_contact(messages_list):
    global contacts
    global owner_id
    for i in range(len(messages_list)):
        if messages_list[i].get("type") == "message":
            temp_id = messages_list[i].get("from_id")
            if temp_id is None:
                print("smth went wrong registering contact, message: ", messages_list[i])
            if temp_id == owner_id:
                continue
            if temp_id not in contacts:
                name = messages_list[i].get("from")
                if name is None:
                    print("smth went wrong registering contact, message: ", messages_list[i])
                init_contact(temp_id, name)
                return temp_id
            else:
                print("looks like id ", temp_id, "is already registered")


def analyze_message(message, sender):
    sender["message_counter"] += 1
    words = message.split()
    for word in words:
        sender["symbol_counter"] += len(word)
        sender["word_counter"] += 1
        for symbol in word:
            if symbol in sender["symb_usage"]:
                sender["symb_usage"][symbol] += 1
            else:
                sender["symb_usage"][symbol] = 1
        if(word[-1] == ','):
            sender['comma_ctr'] += 1
        temp = word.replace(',', '').lower()
        if temp in sender['words_usage']:
            sender['words_usage'][temp] += 1
        else:
            sender['words_usage'][temp] = 1

def compare_time(date):
    global last_time_sent
    if date['day'] == last_time_sent['day']:
        if date['hour'] == last_time_sent['hour']:
            pass
            #minutes
        else:
            pass


def analyze_time(time_sent, sender):
    global last_sender
    time_arr = time_sent.split('T')
    day = time_arr[0].split('-')
    minute = time_arr[1].split(':')
    map(int, day)
    map(int, minute)
    date = {}
    date['year'], date['month'], date['day'], date['hour'], date['minute'] = day[0], day[1], day[2], minute[0], minute[1]
    print(date)
    if sender != last_sender:
        compare_time(date)
        last_sender = sender


def analyze_message_list(message_list, second_id):
    global contacts
    for i in range (len(message_list)):
        message = messages_list[i]
        if message.get("type") == "message":
            sender_id = message.get("from_id")
            if sender_id is None:
                print ("Something went wrong while analyzing message. Exit")
                sys.exit()
            if sender_id not in contacts:
                print('Smth went wrong, sender not in contacts')
                sys.exit()
            if sender_id == owner_id:
                sender = contacts[owner_id][second_id]
            else:
                sender = contacts[sender_id]

            # getting only year. If it is 1970, message was not edited
            time_edited = message.get('edited')[:4]
            if time_edited != '1970':
                sender['msg_edited'] += 1

            time_sent = message.get('date')
            if time_sent is None:
                print('smth went wrong while getting date of sending message. Exit..')
                sys.exit()
            #analyze_time(time_sent, sender)

            text = message.get('text')
            if isinstance(text, str):
                analyze_message(text, sender)
            elif isinstance(text, dict):
                print("i dunno what's it")
            else:
                for elem in text:
                    if isinstance(elem, str):
                        analyze_message(elem, sender)
                    elif isinstance(elem, dict):
                        if elem.get('type') == 'link':
                            sender['links_shared'] += 1
                        analyze_message(elem.get('text'), sender)
        else:
            pass
    calc_avg_len(contacts[owner_id][second_id])
    calc_avg_len(contacts[second_id])


def calc_avg_len(user):
    user['avg_len'] = user['word_counter']/user['message_counter']


def print_data():
    owner = contacts[owner_id]
    owner_name = " "
    for key, value in owner.items():
        if key == 'name':
            owner_name = value
            print("owner's name: ", owner_name)
            print()
        else:
            owner_data = value
            user_data = contacts[key]
            print('chat with: ', user_data['name'])
            print(owner_name, ":")
            for owner_key, owner_value in owner_data.items():
                if owner_key == 'words_usage':
                    print(most_popular_symb_number, " most popular words: ", owner_value.most_common(most_popular_words_number))
                elif owner_key == 'symb_usage':
                    print(most_popular_symb_number, " most popular letters: ", owner_value.most_common(most_popular_symb_number))
                else:
                    print(owner_key, ": ", owner_value)
            print('---- partner ---- ')
            for user_key, user_value in user_data.items():
                if user_key == 'words_usage':
                    print(most_popular_words_number, " most popular words: ", user_value.most_common(most_popular_words_number))
                elif user_key == 'symb_usage':
                    print(most_popular_symb_number, " most popular letters: ", user_value.most_common(most_popular_symb_number))
                else:
                    print(user_key, ": ", user_value)
            print()


try:
    with open(file_path, "r", encoding="utf8") as file:
        data = json.load(file)


except:
    print("smth went wrong while opening file")
    sys.exit()


else:
    owner_id = get_owner_id(data)
    init_owner_contact(owner_id, get_owner_name(data))
    chats = get_chats(data)
    insertionSort(chats)
    for i in range(chats_to_analyze):
        chat_info = chats[i]
        chat_type = get_chat_type(chat_info)
        messages_list = get_messages(chat_info)
        if chat_type == "saved_messages":
            print("You have ", len(messages_list), " saved messages")
        elif chat_type == "personal_chat":
            total_msg_ctr = len(messages_list)
            second_id = register_contact(messages_list)
            analyze_message_list(messages_list, second_id)
        else:
            print('unexpected chat type: ', chat_type)

finally:
    print()
    #print_data()
    print()
    print("\nReading finished. Closing..")
    file.close()
