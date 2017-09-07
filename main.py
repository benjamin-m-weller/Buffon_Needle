import random
import math

def thow_needle(needle_length, tile_seperation_distance):
    # I'm allowed to have the center placed between 0 and (.5 * tile_seperation_distance)
    # simply because I will have an "infinite" series of tiles to land on,
    # as such the model given in class will always apply.
    center_x = random.uniform(0, .5 * tile_seperation_distance)

    # I'm allowed to stop my random numbers at pi/2 because this would also fall in to the model shown in class.
    # There is no value added by allowing a full 180 degree or pi rotation simply because I could view the model
    # "upside down" in cases where there would be more than 90 degree or pi/2 rotation,
    # and as such I would effectively get the same output.
    theta = random.uniform(0, math.pi / 2)

    # Using cosine here to understand if the needle and the line are touching/crossing.
    tip_x = center_x - (needle_length / 2.0) * math.cos(theta)

    return bool(tip_x < 0)


def throw_needles(number):
    # I know I'm going to run 1000 throws but I want to be sure to get them split up nicely for Excel output
    throws = 0
    hits = 0

    iterations=int((number / 50))

    for iteration in range(int((number / 50))):
        # Excel write function here
        for i in range(50):
            needle_length = 2  # needle 2 inches
            tile_seperation_distance = 2  # cracks 2 inch spacing
            if (thow_needle(needle_length, tile_seperation_distance)):
                hits += 1
            throws += 1
        if (hits==0):
            pi = 0
            print("Iteration: " + str(iteration) + " Hits: " + str(hits) + " Throws: " + str(throws) + " Pi: " + str(pi))
        else:
            pi = 2 * (throws / hits)
            print("Iteration: " + str(iteration) + " Hits: " + str(hits) + " Throws: " + str(throws) + " Pi: " + str(pi))


if __name__=="__main__":
    throw_needles(1000)