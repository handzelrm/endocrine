import pickle

def main():
    with open('/home/pi/Documents/python/sent_last_week.pickle', 'rb') as f:
        sent_last_week = pickle.load(f)

    print('sent_last_week value: {}'.format(sent_last_week))
        

if __name__ == '__main__':
    main()
