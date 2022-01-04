import pickle

with open('label.pkl', 'rb') as f:
    label = pickle.load(f)

print(label)