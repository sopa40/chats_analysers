# count number of first texts and show them

# TODO: Replace new_date and old_date (atm used only for hours difference) with some datetime functions.

# TODO: Replace data finding and probably some other functions with regex

# TODO: Redefine friend_name and my_name by searching them in chat

import datetime
import sys
from collections import Counter

file_name = 'chat.txt'
friend_name = 'Qianli Wang'
my_name = 'Sopiha Nazar'

# for counting messages
frnd_msg_ctr = my_msg_ctr = 0

# for counting first messages
friend_initiative = my_initiative = 0

# counting words
frnd_word_ctr = my_word_ctr = 0
frnd_wordcount = my_wordcount = {}
frnd_cnt = Counter()
my_cnt = Counter()

# days chatted
mon = tue = wed = thu = fri = sat = sun = 0

# to know sender of long mesages
current_person = 'unknown'

# actual message without date and name of sender
message = 'unknown'

# for multiple lines messages
long_msg = 'unknown'

# for time tracking
days_chatting = 0
old_date = first_date = last_date = [0, 0, 0, 0]
new_weekday = old_weekday = 0


def number_to_day(number):
    global mon, tue, wed, thu, fri, sat, sun
    if number == 0:
        mon += 1
    elif number == 1:
        tue += 1
    elif number == 2:
        wed += 1
    elif number == 3:
        thu += 1
    elif number == 4:
        fri += 1
    elif number == 5:
        sat += 1
    elif number == 6:
        sun += 1
    else:
        print('false date input')
        sys.exit()


try:
    with open(file_name, 'r', encoding="utf8") as file:
        lines_list = list(file)

except:
    print('smth went wrong while opening file')
    sys.exit()

else:
    for i in range(100):
        if friend_name in lines_list[i] or my_name in lines_list[i]:
            year = int(lines_list[i][6:8]) + 2000
            month = int(lines_list[i][3:5])
            day = int(lines_list[i][:2])
            hours = int(lines_list[i][10:12])
            first_date = [year, month, day, hours]
            break
    for j in range(1, 100):
        i = -j
        if friend_name in lines_list[i] or my_name in lines_list[i]:
            year = int(lines_list[i][6:8]) + 2000
            month = int(lines_list[i][3:5])
            day = int(lines_list[i][:2])
            hours = int(lines_list[i][10:12])
            last_date = [year, month, day, hours]
            break

    for line in lines_list:
        if friend_name in line or my_name in line:
            if friend_name in line:
                current_person = friend_name
            elif my_name in line:
                current_person = my_name

            year = int(line[6:8]) + 2000
            month = int(line[3:5])
            day = int(line[:2])
            hours = int(line[10:12])
            minutes = int(line[13:15])
            new_date = [year, month, day, hours, minutes]
            date_string = datetime.datetime(year, month, day, hours, minutes)
            new_weekday = date_string.weekday()

            # skipping date and nickname info in line
            message = line.split(": ", 2)[1]

            if new_weekday != old_weekday:
                number_to_day(new_weekday)
                days_chatting += 1

            if new_weekday != old_weekday or abs(new_date[3] - old_date[3]) > 2:
                if current_person == friend_name:
                    friend_initiative += 1
                else:
                    my_initiative += 1

            old_date = new_date
            old_weekday = new_weekday

        else:
            message = line

        words = message.split()
        if current_person == friend_name:
            frnd_msg_ctr += 1
            frnd_word_ctr += len(words)
            for word in words:
                if (word[-1] == ','):
                        word.replace(',', '')
                if ("<Medien" in word or "ausgeschlossen>" in word):
                    continue
                frnd_cnt[word] += 1

        else:
            my_msg_ctr += 1
            my_word_ctr += len(words)
            for word in words:
                if (word[-1] == ','):
                        word.replace(',', '')
                if ("<Medien" in word or "ausgeschlossen>" in word):
                    continue
                my_cnt[word] += 1


finally:
    print("\nfriend message counter: ", frnd_msg_ctr)
    print("friend words counter: ", frnd_word_ctr)
    print("words per msg friend: ", frnd_word_ctr / frnd_msg_ctr)
    print("first messages by a friend: ", friend_initiative)
    print("\nmy message couner: ", my_msg_ctr)
    print("my words counter: ", my_word_ctr)
    print("words per msg I: ", my_word_ctr / my_msg_ctr)
    print("first messages by me: ", my_initiative)
    print("\ndays chatting: ", days_chatting)
    print("\nweekdays chatting: \nmonday - ", mon, "\ntuesday - ", tue, "\nwednesday - ", wed, "\nthursday - ", thu,
          "\nfriday - ", fri, "\nsaturday - ", sat, "\nsunday - ", sun)
    print("\ntotal time passed since 1st message: ", last_date[0] - first_date[0], "year(s) ", last_date[1] - first_date[1], "month(s) ",
          last_date[2] - first_date[2], "day(s)")
    print("\nfriend counter: ", frnd_cnt.most_common(15))
    print("\nmy counter: ", my_cnt.most_common(15))
    print("\nReading finished. Closing..")

    file.close()
