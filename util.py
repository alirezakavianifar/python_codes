import numpy as np
import matplotlib.pyplot as plt


def plotRunningAverage(totalrewards):
    N = len(totalrewards)
    running_avg = np.empty(N)
    for t in range(N):
        running_avg[t] = np.mean(totalrewards[max(0, t - 100):(t + 1)])
    plt.plot(running_avg)
    plt.title("Running Average")
    plt.show()


def plotRunningAverageComparison(Algo1, Algo2, labels=None):
    N_1 = len(Algo1)
    N_2 = len(Algo2)
    runningAvgAlgo1 = np.empty(N_1)
    runningAvgAlgo2 = np.empty(N_2)
    for t in range(N_1):
        runningAvgAlgo1[t] = np.mean(Algo1[max(0, t - 100):(t + 1)])
        runningAvgAlgo2[t] = np.mean(Algo2[max(0, t - 100):(t + 1)])

    plt.plot(runningAvgAlgo1, 'r--')
    plt.plot(runningAvgAlgo2, 'b--')
    plt.title("Running Average")
    if labels:
        plt.legend((labels[0], labels[1]))
    plt.show()


def plot_learning_curve(x, scores, figure_file):
    running_avg = np.zeros(len(scores))
    for i in range(len(running_avg)):
        running_avg[i] = np.mean(scores[max(0, i - 100):(i + 1)])
    plt.plot(x, running_avg)
    plt.title('Running average of previous 100 scores')
    plt.savefig(figure_file)


def plot_learning(x, scores, epsilons, filename):
    fig = plt.figure()
    ax = fig.add_subplot(111, label="1")
    ax2 = fig.add_subplot(111, label="2", frame_on=False)

    ax.plot(x, epsilons, color="C0")
    ax.set_xlabel("Training Steps", color="C0")
    ax.set_ylabel("Epsilon", color="C0")
    ax.tick_params(axis='x', colors="C0")
    ax.tick_params(axis='y', colors="C0")

    N = len(scores)
    running_avg = np.empty(N)
    for t in range(N):
        running_avg[t] = np.mean(scores[max(0, t - 100):(t + 1)])

    ax2.scatter(x, running_avg, color="C1")
    ax2.axes.get_xaxis().set_visible(False)
    ax2.yaxis.tick_right()
    ax2.set_ylabel('Score', color="C1")
    ax2.yaxis.set_label_position('right')
    ax2.tick_params(axis='y', colors="C1")

    plt.savefig(filename)
