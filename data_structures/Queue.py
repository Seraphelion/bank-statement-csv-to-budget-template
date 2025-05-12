class Node:
    # Initialization for a node [O(1)]
    def __init__(self, value = None):
        self.value = value
        self.next = None

class Queue:
    # Initialization for a queue [O(1)]
    def __init__(self, value = None):
        self.root = Node(value)
        self.head = self.root
        self.tail = self.root
        self.size = 0
        if value is not None:
            self.size += 1

    # Make the values printable strings [O(1)]
    def __str__(self):
        return f"{self.root.value}"
    
    # Enables len() function [O(1)]
    def __len__(self):
        return self.size
    
    # Adds value at the end of the queue [O(1)]
    def enqueue(self, value):
        if self.root.value is None: # If initial value is None
            self.root.value = value
            self.size += 1
            return True
        
        self.tail.next = Node(value) # Else add node
        self.tail = self.tail.next
        self.size += 1
        return True
    
    # Returns value of first element and removes it [O(1)]
    def dequeue(self):
        if self.isEmpty():
            print("Empty!")
            return None

        temp = self.head.value
        if self.root.next is None:
            self.root.value = None
        else:
            self.root = self.root.next
            self.head = self.root
        self.size -= 1
        return temp

    # Returns queue as an a list [O(n)]
    def toList(self):
        list = []
        root = self.root
        while root is not None:
            list.append(root.value)
            root = root.next
        return list

    # Checks if Queue is empty [O(1)]
    def isEmpty(self):
        if self.root.value is None:
            return True
        return False

    # Prints the queue in the terminal [O(n)]
    def print(self):
        print(self.toList())
        return True

    # Returns value of first element (without removing it) [O(1)]
    def getHead(self): 
        return self.head.value

    # Returns value of last element (without removing it) [O(1)]
    def getTail(self):
        return self.tail.value

