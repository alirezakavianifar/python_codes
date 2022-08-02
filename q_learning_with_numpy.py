import numpy as np
import gym
import matplotlib.pyplot as plt

pos_space = np.linspace(-1.2, 0.6, 20)
vel_space = np.linspace(-0.07, 0.07, 20)

def get_state(observation):
    pos, vel = observation
    pos_bin = np.digitize(pos, pos_space)
    vel_bin = np.digitize(vel, vel_space)
    
    return (pos_bin, vel_bin)

if __name__ == '__main__':
    env = gym.make('MountainCar-v0')
    n_games = 1
    
    for i in range(n_games):
        done = False
        obs = env.reset()
        state = get_state(obs)
        print(obs, state)
