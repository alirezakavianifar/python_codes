from gym import Env
from gym.spaces import Discrete, Box
import numpy as np
import random
import numpy as np
import pandas as pd


class CloudEnv(Env):
    def __init__(self):
        # Actions we can take, down, stay, up
        df = pd.read_excel(r"D:\projects\data\cloudrm.xlsx")
        self.action_space = Discrete(334)
        # Temprature array
        self.observation_space = Box(low=np.array([4422, 500]), high=np.array([9111, 1000]), dtype=np.float32)
        # Set start temp
        self.state = self.observation_space.sample()
        # Set shower length
        self.time_length = 150

    def step(self, action):
        # Apply action
        # Reduce shower length by 1 second
        self.time_length -= 1

        # Calculate reward
        reward = -((0.8 * self.state[0]) + (0.2 * self.state[1]))

        # Check if shower is done
        if self.time_length <= 0:
            done = True

        else:
            done = False
        info = {}

        # Return step information
        return self.state, reward, done, info

    def render(self):
        pass

    def reset(self):
        # Reset shower tempreture
        self.state = self.observation_space.sample()
        # Reset shower time
        self.time_length = 150
        return self.state


