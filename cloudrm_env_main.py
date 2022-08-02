from deep_qlearning_with_keras import Agent
import numpy as np
from util import plot_learning
import gym
from cloudrm_env import CloudEnv

if __name__ == '__main__':
    env = CloudEnv()
    n_games = 500
    agent = Agent(gamma=0.99, epsilon=1.0, alpha=0.0005, input_dims=2,
                  n_actions=344, mem_size=100000, batch_size=64, epsilon_end=0.01)

    scores = []
    eps_history = []

    for i in range(n_games):
        done = False
        score = 0
        observation = env.reset()
        while not done:
            action = agent.choose_action(observation)
            observation_, reward, done, info = env.step(action)
            score += reward
            agent.remember(observation, action, reward, observation_, done)
            observation = observation_
            agent.learn()

        eps_history.append(agent.epsilon)
        scores.append(score)

        avg_score = np.mean(scores[max(0, i-100): (i+1)])
        print('episode ', i, 'score %.2f' % score,
              'average score %.2f' % avg_score)

        if i % 10 == 0 and i > 0:
            agent.save_model()
    filename = 'LunarLander.png'
    x = [i+1 for i in range(n_games)]
    plot_learning(x, scores, eps_history, filename)


