class Node:
    def __init__(self, value = None):
        self.value = value
        self.next = None

class LinkedList:
    def __init__(self, value = None):
        new_node = Node(value)
        self.head = new_node
        self.tail = new_node
        self.lenght = 1

    def __iter__(self):
        current = self.head
        while current is not None:
            yield current
            current = current.next
    
    def __next__(self):
        if self.head is None:
            raise StopIteration
        else:
            current = self.head
            self.head = self.head.next
            return current.value
        
    def __getitem__(self, index):
        if isinstance(index, slice):
            # Handle slicing
            start, stop, step = index.indices(self.lenght)
            result = []
            current = self.head
            for i in range(start, stop + 1, step):
                if current is None:
                    break
                result.append(current.value)
                current = current.next
            return result
        elif isinstance(index, int):
            # Handle int
            if index < 0 or index >= self.lenght:
                return None
            current = self.head
            for _ in range(index):
                current = current.next
            return current.value
        else:
            raise TypeError("Index must be an integer or a slice")

    def append(self, value):
        new_node = Node(value)
        if self.head is None:
            self.head = new_node
            self.tail = new_node
        else:
            self.tail.next = new_node
            self.tail = new_node
        self.lenght += 1
        return True
    
    def prepend(self, value):
        new_node = Node(value)
        if self.head is None:
            self.head = new_node
            self.tail = new_node
        else:
            pre = new_node
            pre.next = self.head
            self.head = pre
        self.lenght += 1
        return True

    def pop(self):
        if self.lenght == 0:
            return None
        
        temp = self.head
        pre = self.head

        if self.lenght == 1:
            self.head = None
            self.tail = None
            self.lenght -= 1
            return temp.value

        while(temp.next):
            pre = temp
            temp = temp.next
        self.tail = pre
        self.tail.next = None
        self.lenght -= 1

    def size(self):
        return self.lenght

    def isEmpty(self):
        if self.size() == 0:
            return True
        else:
            return False
        
    def print_list(self):
        temp = self.head
        while temp is not None:
            print(temp.value)
            temp = temp.next

    def contains(self, value):
        if self.lenght == 0:
            return False
        
        temp = self.head
        while temp is not None:
            if temp.value == value:
                return True
            temp = temp.next