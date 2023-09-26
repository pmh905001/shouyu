from queue import Queue


class TaskExecutor:

    def __init__(self):
        self.q = Queue(20)

    def add(self, fun, args):
        self.q.put((fun, args))

    def run(self):
        while True:
            msg = self.q.get()
            fun = msg[0]
            fun(*msg[1])
            # time.sleep(1)
            self.q.task_done()
